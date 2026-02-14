/**
 * ユーティリティ関数
 */

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

/* ---------- Generic Sheet Helpers ---------- */
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