/**
 * API・コントローラー
 */

function getMemberData(year, month) {
  return _fetchDataWithCache(year, month);
}

function getAdminData(year, month) {
  return _fetchDataWithCache(year, month);
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

  _clearCache(year, month); // 保存時にキャッシュクリア
  return {
    ok: true,
    schedule: getScheduleForMonth(year, month)
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
  _clearCache(year, month); // キャッシュクリア
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
  _clearCache(year, month); // キャッシュクリア

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

  _clearCache(year, month); // キャッシュクリア
  return { ok: true, message: "保存しました。", days: result };
}