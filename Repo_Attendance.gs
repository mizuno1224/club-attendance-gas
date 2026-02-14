/**
 * 出欠シート操作
 */

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

// 1日分だけ書き込み（Upsert）— 1回の setValues でまとめて書き込み
function _writeAttendanceDay(year, month, day, presentList, absentList, tardyList, earlyList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = _getAttendanceSheet();
  if (!sh) sh = ss.insertSheet(ATTENDANCE_SHEET_CANDIDATES[0]);
  _ensureAttendanceHeader(sh);
  const vals = sh.getDataRange().getValues();
  const map = _detectAttendanceHeader(vals[0]);
  const dateObj = new Date(year, month - 1, day);
  const presentStr = _uniq(presentList || []).join(",");
  const absentStr = _uniq(absentList || []).join(",");
  const tardyStr = _uniq(tardyList || []).join(",");
  const earlyStr = _uniq(earlyList || []).join(",");

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
    rowIdx = vals.length + 1;
    const newRow = new Array(vals[0].length).fill("");
    newRow[map.date] = dateObj;
    if (map.present >= 0) newRow[map.present] = presentStr;
    if (map.absent >= 0) newRow[map.absent] = absentStr;
    if (map.tardy >= 0) newRow[map.tardy] = tardyStr;
    if (map.early >= 0) newRow[map.early] = earlyStr;
    vals.push(newRow);
    sh.getRange(vals.length, 1, vals.length, newRow.length).setValues([newRow]);
    sh.getRange(vals.length, map.date + 1).setNumberFormat("yyyy/MM/dd");
    return;
  }

  const r = rowIdx - 1;
  const row = vals[r].slice();
  row[map.date] = dateObj;
  if (map.present >= 0) row[map.present] = presentStr;
  if (map.absent >= 0) row[map.absent] = absentStr;
  if (map.tardy >= 0) row[map.tardy] = tardyStr;
  if (map.early >= 0) row[map.early] = earlyStr;
  sh.getRange(rowIdx, 1, rowIdx, row.length).setValues([row]);
  sh.getRange(rowIdx, map.date + 1).setNumberFormat("yyyy/MM/dd");
}

// 月全体を書き戻す（saveMemberResponse 用）— 一括読み書きで高速化
function _writeAttendanceToSheet(year, month, attendanceMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = _getAttendanceSheet();
  if (!sh) sh = ss.insertSheet(ATTENDANCE_SHEET_CANDIDATES[0]);
  _ensureAttendanceHeader(sh);

  const range = sh.getDataRange();
  const vals = range.getValues();
  const map = _detectAttendanceHeader(vals[0]);
  if (map.date < 0) return;

  const dateToRowIdx = {};
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r][map.date];
    if (!v) continue;
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d)) continue;
    if (d.getFullYear() === year && (d.getMonth() + 1) === month) {
      dateToRowIdx[d.getDate()] = r;
    }
  }

  const days = Object.keys(attendanceMap || {});
  days.forEach(k => {
    const day = parseInt(k, 10);
    if (isNaN(day)) return;
    const A = attendanceMap[k] || {};
    const present = _uniq([...(A.morning || []), ...(A.afternoon || []), ...(A.after || [])]);
    const absent = _uniq(A.absent || []);
    const tardy = _uniq(A.tardy || []);
    const early = _uniq(A.early || []);

    let r = dateToRowIdx[day];
    if (r === undefined) {
      const newRow = new Array(vals[0].length).fill("");
      newRow[map.date] = new Date(year, month - 1, day);
      vals.push(newRow);
      r = vals.length - 1;
      dateToRowIdx[day] = r;
    }

    if (map.present >= 0) vals[r][map.present] = present.join(",");
    if (map.absent >= 0) vals[r][map.absent] = absent.join(",");
    if (map.tardy >= 0) vals[r][map.tardy] = tardy.join(",");
    if (map.early >= 0) vals[r][map.early] = early.join(",");
  });

  if (vals.length > 0) {
    sh.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
    if (vals.length > 1) {
      sh.getRange(2, map.date + 1, vals.length, map.date + 1).setNumberFormat("yyyy/MM/dd");
    }
  }
}