/** ============================================================
 * 女子軟式野球部 出欠管理（出欠はシート保存 / 手動登録 / 祝日 / 1日単位保存API）
 * schedule[day]  = { morning:boolean, afternoon:boolean, after:boolean, note:string, place:string, time:string }
 * attendance[day]= { morning:[], afternoon:[], after:[], absent:[], tardy:[], early:[] } // presentは3枠に複製
 * ============================================================ */

const DEFAULT_ROSTER = ["ゆうり", "みゆ", "のぞみ", "えみり", "まな", "まほ", "まい", "しん"];
const TZ = "Asia/Tokyo";
const PROP_LOG_SHEET_ID = "LOG_SHEET_ID";

const SCHEDULE_SHEET_CANDIDATES = ["活動予定", "スケジュール", "Schedule", "schedule"];
const ATTENDANCE_SHEET_CANDIDATES = ["出欠", "出席", "Attendance", "attendance"];

const HEADER_ALIASES = {
  // schedule
  date: ["日付", "日", "Date", "date", "DATE", "日にち"],
  off: ["休み", "オフ", "OFF", "off", "Off", "中止"],
  morning: ["午前", "AM", "am", "朝", "morning", "Morning"],
  afternoon: ["午後", "PM", "pm", "afternoon", "Afternoon"],
  after: ["業後", "after", "After", "afterWork", "夜", "夜間", "ナイター"],
  note: ["メモ", "備考", "内容", "note", "Note"],
  place: ["場所", "会場", "place", "Place", "グラウンド"],
  time: ["時間", "活動時間", "集合", "集合時間", "gather", "gatherAt", "time", "Time"],
  // attendance
  present: ["出席", "参加", "present", "Present"],
  absent: ["欠席", "不参加", "absent", "Absent"],
  tardy: ["遅刻", "tardy", "Tardy", "Late", "遅参"],
  early: ["早退", "early", "Early", "早上がり"]
};

function _props() {
  return PropertiesService.getScriptProperties();
}

function _zero(n) {
  return ('0' + n).slice(-2);
}

function _ymd(d) {
  return d.getFullYear() + "-" + _zero(d.getMonth() + 1) + "-" + _zero(d.getDate());
}

function _toBool(v) {
  if (v === true) return true;
  if (typeof v === "number") return v !== 0;
  if (v instanceof Date) return true;
  if (typeof v === "string") {
    const s = v.trim().toLowerCase();
    return ["true", "1", "on", "yes", "y", "ok", "有", "あり", "実施", "○", "◯"].some(k => s.indexOf(k) !== -1);
  }
  return false;
}

function _toStr(v) {
  return (v == null) ? "" : String(v);
}

function _splitNames(v) {
  return _toStr(v).split(/[,"\s、]+/).map(s => s.trim()).filter(Boolean);
}

function _uniq(arr) {
  return Array.from(new Set(arr));
}

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

/* ---------- Schedule sheet ---------- */
function _getScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of SCHEDULE_SHEET_CANDIDATES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function _detectHeaderMap(row, mapKeys) {
  const find = (aliases) => {
    for (let c = 0; c < row.length; c++) {
      const t = (row[c] || "").toString().trim();
      if (!t) continue;
      for (const a of aliases) {
        if (t.toLowerCase() === a.toLowerCase()) return c;
      }
    }
    return -1;
  };
  const m = {};
  Object.keys(mapKeys).forEach(k => m[k] = find(mapKeys[k]));
  return m;
}

function _ensureHeader(sh, headers) {
  const first = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
  const row = [];
  for (let i = 0; i < headers.length; i++) {
    row[i] = first[i] ? first[i] : headers[i];
  }
  sh.getRange(1, 1, 1, row.length).setValues([row]);
}

function _readScheduleFromSheet(year, month, optValues) {
  const sh = _getScheduleSheet();
  if (!sh) return {};
  const values = optValues || sh.getDataRange().getValues();
  if (values.length < 2) return {};
  const map = _detectHeaderMap(values[0], {
    date: HEADER_ALIASES.date,
    off: HEADER_ALIASES.off,
    morning: HEADER_ALIASES.morning,
    afternoon: HEADER_ALIASES.afternoon,
    after: HEADER_ALIASES.after,
    note: HEADER_ALIASES.note,
    place: HEADER_ALIASES.place,
    time: HEADER_ALIASES.time
  });
  if (map.date < 0) return {};
  const out = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rawDate = row[map.date];
    if (!rawDate) continue;
    let d;
    if (rawDate instanceof Date) d = new Date(rawDate.getTime());
    else if (typeof rawDate === "number") d = new Date(rawDate);
    else {
      const p = new Date(rawDate);
      if (isNaN(p)) continue;
      d = p;
    }
    if (d.getFullYear() !== year || (d.getMonth() + 1) !== month) continue;
    const day = d.getDate();
    out[day] = {
      off: map.off >= 0 ? _toBool(row[map.off]) : false,
      morning: map.morning >= 0 ? _toBool(row[map.morning]) : false,
      afternoon: map.afternoon >= 0 ? _toBool(row[map.afternoon]) : false,
      after: map.after >= 0 ? _toBool(row[map.after]) : false,
      note: map.note >= 0 ? _toStr(row[map.note]) : "",
      place: map.place >= 0 ? _toStr(row[map.place]) : "",
      time: map.time >= 0 ? _toStr(row[map.time]) : ""
    };
  }
  return out;
}

function _writeScheduleToSheet(year, month, patch) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = _getScheduleSheet();
  if (!sh) sh = ss.insertSheet(SCHEDULE_SHEET_CANDIDATES[0]);
  const HEADERS = ["日付", "午前", "午後", "業後", "メモ", "場所", "時間"];
  _ensureHeader(sh, HEADERS);
  // 最新のデータを全取得
  const range = sh.getDataRange();
  const vals = range.getValues();
  // ヘッダー再マッピング
  const map = _detectHeaderMap(vals[0], {
    date: HEADER_ALIASES.date,
    off: HEADER_ALIASES.off,
    morning: HEADER_ALIASES.morning,
    afternoon: HEADER_ALIASES.afternoon,
    after: HEADER_ALIASES.after,
    note: HEADER_ALIASES.note,
    place: HEADER_ALIASES.place,
    time: HEADER_ALIASES.time
  });
  // 足りない列があれば追加
  const ensureCol = (key, title) => {
    if (map[key] >= 0) return;
    const lc = sh.getLastColumn();
    sh.getRange(1, lc + 1).setValue(title);
    vals.forEach(row => row.push("")); // 値配列にも列を追加しておく
    map[key] = lc;
  };
  ensureCol("date", "日付");
  ensureCol("off", "休み");
  ensureCol("morning", "午前");
  ensureCol("afternoon", "午後");
  ensureCol("after", "業後");
  ensureCol("note", "メモ");
  ensureCol("place", "場所");
  ensureCol("time", "時間");

  // 日付 -> 行インデックスのマップ作成
  const dateToRowIdx = {};
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r][map.date];
    if (!v) continue;
    let d;
    if (v instanceof Date) d = v;
    else if (typeof v === "number") d = new Date(v);
    else {
      const p = new Date(v);
      if (isNaN(p)) continue;
      d = p;
    }
    dateToRowIdx[_ymd(new Date(d.getFullYear(), d.getMonth(), d.getDate()))] = r;
  }

  // patchの内容をメモリ上の vals に反映
  Object.keys(patch).forEach(k => {
    const day = parseInt(k, 10);
    const dateObj = new Date(year, month - 1, day);
    const key = _ymd(dateObj);
    let r = dateToRowIdx[key];

    // 行が存在しなければ新規作成して追加
    if (r === undefined) {
      const newRow = new Array(vals[0].length).fill("");
      newRow[map.date] = dateObj;
      vals.push(newRow);
      r = vals.length - 1;
      dateToRowIdx[key] = r;
    }

    const p = patch[k] || {};
    if (p.hasOwnProperty("off")) vals[r][map.off] = !!p.off;
    if (p.hasOwnProperty("morning")) vals[r][map.morning] = !!p.morning;
    if (p.hasOwnProperty("afternoon")) vals[r][map.afternoon] = !!p.afternoon;
    if (p.hasOwnProperty("after")) vals[r][map.after] = !!p.after;
    if (p.hasOwnProperty("note")) vals[r][map.note] = _toStr(p.note);
    if (p.hasOwnProperty("place")) vals[r][map.place] = _toStr(p.place);
    if (p.hasOwnProperty("time")) vals[r][map.time] = _toStr(p.time);
  });

  // 一括書き込み
  if (vals.length > 0) {
    sh.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
  }
}

/* ---------- Attendance sheet ---------- */
function _getAttendanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of ATTENDANCE_SHEET_CANDIDATES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function _ensureAttendanceHeader(sh) {
  const first = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  const have = new Map();
  for (let c = 0; c < first.length; c++) {
    const t = (first[c] || "").toString().trim();
    if (t) have.set(t, c);
  }
  const want = ["日付", "出席", "欠席", "遅刻", "早退"];
  want.forEach(title => {
    if (!have.has(title)) {
      const lc = sh.getLastColumn();
      sh.getRange(1, lc + 1).setValue(title);
      have.set(title, lc);
    }
  });
}

function _detectAttendanceHeader(row) {
  return _detectHeaderMap(row, {
    date: HEADER_ALIASES.date,
    present: HEADER_ALIASES.present,
    absent: HEADER_ALIASES.absent,
    tardy: HEADER_ALIASES.tardy,
    early: HEADER_ALIASES.early
  });
}

// 月全体読み込み（UI用）
function _readAttendanceFromSheet(year, month, optValues) {
  let sh = _getAttendanceSheet();
  if (!sh) return {};
  const vals = optValues || sh.getDataRange().getValues();
  if (vals.length < 2) return {};
  const map = _detectAttendanceHeader(vals[0]);
  if (map.date < 0) return {};
  const out = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const rawDate = row[map.date];
    if (!rawDate) continue;
    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (isNaN(d)) continue;
    if (d.getFullYear() !== year || (d.getMonth() + 1) !== month) continue;
    const day = d.getDate();
    const present = _uniq(_splitNames(map.present >= 0 ? row[map.present] : ""));
    const absent = _uniq(_splitNames(map.absent >= 0 ? row[map.absent] : ""));
    const tardy = _uniq(_splitNames(map.tardy >= 0 ? row[map.tardy] : ""));
    const early = _uniq(_splitNames(map.early >= 0 ? row[map.early] : ""));
    out[day] = {
      morning: present.slice(),
      afternoon: present.slice(),
      after: present.slice(),
      absent: absent,
      tardy: tardy,
      early: early
    };
  }
  return out;
}

// 1日分だけ読み込み
function _readAttendanceDay(year, month, day) {
  const sh = _getAttendanceSheet();
  if (!sh) return {
    present: [],
    absent: [],
    tardy: [],
    early: []
  };
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return {
    present: [],
    absent: [],
    tardy: [],
    early: []
  };
  const map = _detectAttendanceHeader(vals[0]);
  if (map.date < 0) return {
    present: [],
    absent: [],
    tardy: [],
    early: []
  };
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r][map.date];
    if (!v) continue;
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d)) continue;
    if (d.getFullYear() === year && (d.getMonth() + 1) === month && d.getDate() === day) {
      return {
        present: _uniq(_splitNames(map.present >= 0 ? vals[r][map.present] : "")),
        absent: _uniq(_splitNames(map.absent >= 0 ? vals[r][map.absent] : "")),
        tardy: _uniq(_splitNames(map.tardy >= 0 ? vals[r][map.tardy] : "")),
        early: _uniq(_splitNames(map.early >= 0 ? vals[r][map.early] : ""))
      };
    }
  }
  return {
    present: [],
    absent: [],
    tardy: [],
    early: []
  };
}

// 1日分だけ書き込み（Upsert）
function _writeAttendanceDay(year, month, day, presentList, absentList, tardyList, earlyList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = _getAttendanceSheet();
  if (!sh) sh = ss.insertSheet(ATTENDANCE_SHEET_CANDIDATES[0]);
  _ensureAttendanceHeader(sh);
  const vals = sh.getDataRange().getValues();
  const map = _detectAttendanceHeader(vals[0]);
  const dateObj = new Date(year, month - 1, day);
  let rowIdx = -1;
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r][map.date];
    if (!v) continue;
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d)) continue;
    if (d.getFullYear() === year && (d.getMonth() + 1) === month && d.getDate() === day) {
      rowIdx = r + 1;
      break;
    }
  }
  if (rowIdx < 0) {
    rowIdx = sh.getLastRow() + 1;
  }
  sh.getRange(rowIdx, map.date + 1).setValue(dateObj);
  if (map.present >= 0) sh.getRange(rowIdx, map.present + 1).setValue(_uniq(presentList || []).join(","));
  if (map.absent >= 0) sh.getRange(rowIdx, map.absent + 1).setValue(_uniq(absentList || []).join(","));
  if (map.tardy >= 0) sh.getRange(rowIdx, map.tardy + 1).setValue(_uniq(tardyList || []).join(","));
  if (map.early >= 0) sh.getRange(rowIdx, map.early + 1).setValue(_uniq(earlyList || []).join(","));
  sh.getRange(rowIdx, map.date + 1).setNumberFormat("yyyy/MM/dd");
}

// 月全体を書き戻す（saveMemberResponse 用）
function _writeAttendanceToSheet(year, month, attendanceMap) {
  const days = Object.keys(attendanceMap || {});
  days.forEach(k => {
    const d = parseInt(k, 10);
    if (isNaN(d)) return;
    const A = attendanceMap[k] || {};
    const present = _uniq([...(A.morning || []), ...(A.afternoon || []), ...(A.after || [])]);
    const absent = _uniq(A.absent || []);
    const tardy = _uniq(A.tardy || []);
    const early = _uniq(A.early || []);
    _writeAttendanceDay(year, month, d, present, absent, tardy, early);
  });
}

/* ---------- Holidays (Japan) ---------- */
function _getHolidaysMap(year, month) {
  const key = "HOLIDAY_CACHE_" + year + "_" + month;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  try {
    const calId = 'ja.japanese#holiday@group.v.calendar.google.com';
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

/* ---------- Graph Data API ---------- */
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

    const rate = totalDays === 0 ? 0 : Math.round((presentCount / totalDays) * 1000) / 10;
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

function getStatsForGraph(name, year, month, count) {
  return getPersonalStats(name, year, month, count);
}

function saveSchedulePatch(year, month, patch, actor) {
  _writeScheduleToSheet(year, month, patch);

  // ログを配列にまとめて一括保存（高速化）
  const logs = [];
  Object.keys(patch || {}).forEach(d => {
    Object.keys(patch[d] || {}).forEach(f => {
      logs.push([actor || "admin", year, month, d, f, "", patch[d][f]]);
    });
  });
  if (logs.length) _appendLogBatch(logs);

  return {
    ok: true,
    schedule: getScheduleForMonth(year, month)
  };
}

function getMemberData(year, month) {
  return {
    schedule: getScheduleForMonth(year, month),
    attendance: getAttendanceForMonth(year, month),
    roster: _getRoster(),
    holidays: _getHolidaysMap(year, month)
  };
}

function getAdminData(year, month) {
  return {
    schedule: getScheduleForMonth(year, month),
    attendance: getAttendanceForMonth(year, month),
    roster: _getRoster(),
    holidays: _getHolidaysMap(year, month)
  };
}

/** 既存：月まとめで更新（互換・遅刻/早退は空で運用するなら未指定でOK） */
function saveMemberResponse(name, year, month, responses) {
  if (!name) return "名前が空です。";
  const attendance = getAttendanceForMonth(year, month);
  // 全日から name を除去 → 渡された日だけ入れ直す
  Object.keys(attendance).forEach(d => {
    ["morning", "afternoon", "after", "absent", "tardy", "early"].forEach(k => {
      attendance[d][k] = (attendance[d][k] || []).filter(n => n !== name);
    });
  });
  Object.keys(responses || {}).forEach(d => {
    const times = responses[d] || [];
    if (!attendance[d]) attendance[d] = {
      morning: [],
      afternoon: [],
      after: [],
      absent: [],
      tardy: [],
      early: []
    };
    if (times.includes("absent")) attendance[d].absent.push(name);
    else {
      if (times.includes("morning")) attendance[d].morning.push(name);
      if (times.includes("afternoon")) attendance[d].afternoon.push(name);
      if (times.includes("after")) attendance[d].after.push(name);
    }
  });
  _writeAttendanceToSheet(year, month, attendance);
  return "送信しました。";
}

/** 新規：1日だけを安全更新（部員ページの自動保存/一括送信が使用） */
function saveMemberResponseDay(name, year, month, day, times) {
  if (!name) return { ok: false, message: "名前が空です。" };
  const cur = _readAttendanceDay(year, month, day); // {present:[],absent:[],tardy:[],early:[]}
  let present = (cur.present || []).filter(n => n !== name);
  let absent = (cur.absent || []).filter(n => n !== name);
  let tardy = (cur.tardy || []).filter(n => n !== name);
  let early = (cur.early || []).filter(n => n !== name);

  let isPresent = false, isAbsent = false;
  let setTardy = false, setEarly = false;

  if (times && times.length) {
    if (times.includes("absent")) isAbsent = true;
    else isPresent = true;
    setTardy = times.includes("tardy");
    setEarly = times.includes("early");
  }

  if (isAbsent) {
    absent.push(name);
  } else if (isPresent) {
    present.push(name);
    if (setTardy) tardy.push(name);
    if (setEarly) early.push(name);
  }

  _writeAttendanceDay(year, month, day, present, absent, tardy, early);

  const dayData = {
    morning: isPresent ? present.slice() : [],
    afternoon: isPresent ? present.slice() : [],
    after: isPresent ? present.slice() : [],
    absent: absent.slice(),
    tardy: tardy.slice(),
    early: early.slice()
  };
  return { ok: true, message: "保存しました。", day: String(day), dayData };
}

/** まとめ送信（部員UIの一括送信が使用） - 高速化版 */
function saveMemberResponseBatch(name, year, month, changes) {
  if (!name) return { ok: false, message: "名前が空です。" };
  if (!changes || !changes.length) return { ok: false, message: "変更がありません。" };

  // 1. シート全体を一度だけ読み込む
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = _getAttendanceSheet();
  if (!sh) sh = ss.insertSheet(ATTENDANCE_SHEET_CANDIDATES[0]);
  _ensureAttendanceHeader(sh);

  const range = sh.getDataRange();
  const vals = range.getValues();
  const map = _detectAttendanceHeader(vals[0]);

  if (map.date < 0) return { ok: false, message: "日付列が見つかりません。" };
  // 2. 日付 -> 行インデックスのマップを作成
  const dateToRowIdx = {};
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r][map.date];
    if (!v) continue;
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d)) continue;
    // 対象の年月のみマッピング
    if (d.getFullYear() === year && (d.getMonth() + 1) === month) {
      dateToRowIdx[d.getDate()] = r;
    }
  }

  const result = {};

  // 3. メモリ上で変更を適用
  changes.forEach(item => {
    const day = parseInt(item.day, 10);
    if (isNaN(day)) return;

    let r = dateToRowIdx[day];

    // 行がなければ新規作成（メモリ上）
    if (r === undefined) {
      const newRow = new Array(vals[0].length).fill("");
      newRow[map.date] = new Date(year, month - 1, day);
      vals.push(newRow);
      r = vals.length - 1;
      dateToRowIdx[day] = r;
    }

    // 現在のセルの値を取得・パース
    const getArr = (idx) => (idx >= 0 && vals[r][idx]) ? _splitNames(vals[r][idx]) : [];

    let present = getArr(map.present).filter(n => n !== name);
    let absent = getArr(map.absent).filter(n => n !== name);
    let tardy = getArr(map.tardy).filter(n => n !== name);
    let early = getArr(map.early).filter(n => n !== name);

    const times = item.times || [];
    const isAbsent = times.includes("absent");
    const isPresent = !isAbsent && times.length > 0;

    if (isAbsent) {
      absent.push(name);
    } else if (isPresent) {
      present.push(name);
      if (times.includes("tardy")) tardy.push(name);
      if (times.includes("early")) early.push(name);
    }

    // 重複除去してメモリに書き戻し
    const setCell = (idx, arr) => {
      if (idx >= 0) vals[r][idx] = _uniq(arr).join(",");
    };
    setCell(map.present, present);
    setCell(map.absent, absent);
    setCell(map.tardy, tardy);
    setCell(map.early, early);

    result[day] = {
      morning: present.slice(),
      afternoon: present.slice(),
      after: present.slice(),
      absent: absent.slice(),
      tardy: tardy.slice(),
      early: early.slice()
    };
  });

  // 4. 一括でシートに書き込む
  if (vals.length > 0) {
    const destRange = sh.getRange(1, 1, vals.length, vals[0].length);
    destRange.setValues(vals);
    // 日付列のフォーマット統一
    sh.getRange(2, map.date + 1, vals.length - 1, 1).setNumberFormat("yyyy/MM/dd");
  }

  return { ok: true, message: "保存しました。", days: result };
}

/* ---------- Log ---------- */
function _getOrCreateLogSheet() {
  const props = _props();
  let ssId = props.getProperty(PROP_LOG_SHEET_ID);
  let ss;
  if (!ssId) {
    ss = SpreadsheetApp.create("Club Operation Log");
    props.setProperty(PROP_LOG_SHEET_ID, ss.getId());
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }
  let sh = ss.getSheetByName("OperationLog");
  if (!sh) {
    sh = ss.insertSheet("OperationLog");
    sh.appendRow(["Timestamp(JST)", "Actor", "Year", "Month", "Day", "Field", "OldValue", "NewValue"]);
  }
  return sh;
}

function _appendLog(actor, year, month, day, field, oldV, newV) {
  _appendLogBatch([
    [actor, year, month, day, field, oldV, newV]
  ]);
}

function _appendLogBatch(logRows) {
  try {
    const sh = _getOrCreateLogSheet();
    const nowStr = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    const rows = logRows.map(r => [
      nowStr,
      r[0] || "admin", r[1], r[2], r[3], r[4], String(r[5] ?? ""), String(r[6] ?? "")
    ]);
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  } catch (e) {}
}

/* ---------- Web app & menu ---------- */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : "member";
  var fileName = (page === "admin") ? "admin" : "member";
  return HtmlService.createHtmlOutputFromFile(fileName).setTitle("女子軟式野球部").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("部活")
    .addItem("部員ページ（新しいタブ）", "openMember")
    .addItem("管理者ページ（新しいタブ）", "openAdmin")
    .addSeparator()
    .addItem("部員ページ（サイドバー表示）", "openMemberSidebar")
    .addItem("管理者ページ（サイドバー表示）", "openAdminSidebar")
    .addToUi();
}

function _getWebAppUrlOrAlert() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert("先にウェブアプリとしてデプロイしてください（デプロイ→新しいデプロイ→ウェブアプリ）。");
    return null;
  }
  return url;
}

function openMember() {
  var url = _getWebAppUrlOrAlert();
  if (!url) return;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + url + '" target="_blank">部員ページを開く</a></p>'), "部員ページ");
}

function openAdmin() {
  var url = _getWebAppUrlOrAlert();
  if (!url) return;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + url + '?page=admin" target="_blank">管理者ページを開く</a></p>'), "管理者ページ");
}

function openMemberSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('member').setTitle('部員ページ');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openAdminSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('admin').setTitle('管理者ページ');
  SpreadsheetApp.getUi().showSidebar(html);
}