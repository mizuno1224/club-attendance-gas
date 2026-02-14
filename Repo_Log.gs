/**
 * ログ操作
 */

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
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, startRow + rows.length - 1, rows[0].length).setValues(rows);
    }
  } catch (e) {}
}