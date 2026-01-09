/**
 * スケジュールシート操作
 */

function _getScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of SCHEDULE_SHEET_CANDIDATES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
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