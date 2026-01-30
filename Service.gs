/**
 * ビジネスロジック・データサービス
 */

/* ---------- Roster ---------- */
function _getRoster() {
  const raw = _props().getProperty("ROSTER_JSON");
  if (!raw) return DEFAULT_ROSTER.slice();
  try {
    return JSON.parse(raw);
  } catch (e) {
    return DEFAULT_ROSTER.slice();
  }
}

/* ---------- Holidays (Japan) ---------- */
function _getHolidaysMap(year, month) {
  const key = "HOLIDAY_CACHE_" + year + "_" + month;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  try {
    // 修正: 公式の祝日のみを含むカレンダーIDに変更（節分などの行事を除外）
    const calId = 'ja.japanese.official#holiday@group.v.calendar.google.com';
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) return {};
    const start = new Date(year, month - 1, 1, 0, 0, 0);
    const end = new Date(year, month, 1, 0, 0, 0);
    const evs = cal.getEvents(start, end);
    const map = {};
    evs.forEach(e => {
      const dt = e.getAllDayStartDate ? e.getAllDayStartDate() : e.getStartTime();
      const d = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
      if (d.getFullYear() === year && d.getMonth() + 1 === month) {
        map[d.getDate()] = e.getTitle();
      }
    });
    cache.put(key, JSON.stringify(map), 21600); // 6時間キャッシュ
    return map;
  } catch (err) {
    return {};
  }
}

/* ---------- Data Aggregation ---------- */
function getScheduleForMonth(year, month, optValues) {
  const s = _readScheduleFromSheet(year, month, optValues);
  const last = new Date(year, month, 0).getDate();
  for (let d = 1; d <= last; d++) {
    if (!s[d]) s[d] = {
      off: false,
      morning: false,
      afternoon: false,
      after: false,
      note: "",
      place: "",
      time: ""
    };
    const v = s[d];
    s[d] = {
      off: !!v.off,
      morning: !!v.morning,
      afternoon: !!v.afternoon,
      after: !!v.after,
      note: _toStr(v.note),
      place: _toStr(v.place),
      time: _toStr(v.time)
    };
  }
  return s;
}

function getAttendanceForMonth(year, month, optValues) {
  return _readAttendanceFromSheet(year, month, optValues);
}

// 共通データ取得関数（キャッシュ対応）
function _fetchDataWithCache(year, month) {
  // 静的データ（スケジュール・名簿・祝日）のキャッシュキー
  const staticKey = "STATIC_" + year + "_" + month; 
  // 動的データ（出席）のキャッシュキー
  const attendanceKey = "ATTENDANCE_" + year + "_" + month;
  
  const cache = CacheService.getScriptCache();
  
  let scheduleData, rosterData, holidaysData, attendanceData;

  // 1. 静的データの取得
  const staticCached = cache.get(staticKey);
  if (staticCached) {
    try {
      const parsed = JSON.parse(staticCached);
      scheduleData = parsed.schedule;
      rosterData = parsed.roster;
      holidaysData = parsed.holidays;
    } catch (e) {}
  }
  
  if (!scheduleData) {
    scheduleData = getScheduleForMonth(year, month);
    rosterData = _getRoster();
    holidaysData = _getHolidaysMap(year, month);
    try {
      cache.put(staticKey, JSON.stringify({
        schedule: scheduleData,
        roster: rosterData,
        holidays: holidaysData
      }), 21600); // 6時間
    } catch (e) {}
  }

  // 2. 出席データの取得
  const attCached = cache.get(attendanceKey);
  if (attCached) {
    try {
      attendanceData = JSON.parse(attCached);
    } catch (e) {}
  }
  
  if (!attendanceData) {
    attendanceData = getAttendanceForMonth(year, month);
    try {
      cache.put(attendanceKey, JSON.stringify(attendanceData), 21600);
    } catch (e) {}
  }

  return {
    schedule: scheduleData,
    roster: rosterData,
    holidays: holidaysData,
    attendance: attendanceData
  };
}

/* ---------- Graph Logic ---------- */
function getPersonalStats(name, startYear, startMonth, count) {
  if (!name) return { months: [], rates: [] };
  const months = [];
  const rates = [];

  // Optimization: Fetch all data once
  const sSh = _getScheduleSheet();
  const sVals = sSh ? sSh.getDataRange().getValues() : [];
  const aSh = _getAttendanceSheet();
  const aVals = aSh ? aSh.getDataRange().getValues() : [];
  let y = startYear, m = startMonth;
  for (let i = 0; i < count; i++) {
    const s = getScheduleForMonth(y, m, sVals);
    const a = getAttendanceForMonth(y, m, aVals);

    let totalDays = 0;
    let presentCount = 0;
    // その月の日数分ループ
    const last = new Date(y, m, 0).getDate();
    for (let d = 1; d <= last; d++) {
      // 活動がある日か判定
      if (s[d] && (s[d].morning || s[d].afternoon || s[d].after)) {
        totalDays++;
        // 出席判定
        const A = a[d] || {};
        const pSet = new Set([...(A.morning || []), ...(A.afternoon || []), ...(A.after || [])]);
        if (pSet.has(name)) presentCount++;
      }
    }

    const rate = totalDays === 0 ?
      0 : Math.round((presentCount / totalDays) * 1000) / 10;
    months.push(`${m}月`);
    rates.push(rate);

    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
  }
  return { months, rates };
}