// CSV 파일을 읽고 파싱하는 유틸리티

export function parseCsv(text) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // BOM 제거
  var rows = [];
  var row = [];
  var field = "";
  var inQuotes = false;
  for (var i = 0; i < text.length; i++) {
    var c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { row.push(field); field = ""; }
      else if (c === '\r') { /* skip */ }
      else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ""; }
      else field += c;
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  rows = rows.filter(function (r) { return !(r.length === 1 && r[0] === ""); });
  if (rows.length === 0) return [];
  var headers = rows[0];
  return rows.slice(1).map(function (r) {
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = r[idx] !== undefined ? r[idx] : ""; });
    return obj;
  });
}

export function readFileAsText(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function () { resolve(reader.result); };
    reader.onerror = reject;
    reader.readAsText(file, "utf-8");
  });
}

export function toFloat(v) {
  if (v === undefined || v === null || v === "") return 0;
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
