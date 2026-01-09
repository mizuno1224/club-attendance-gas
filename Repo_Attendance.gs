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