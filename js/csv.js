// CSV 파일을 읽고 파싱하는 유틸리티

function parseCsv(text) {
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
    obj.__cells = r; // 연차 시간(F/G열)을 헤더명과 무관하게 열 위치로 읽을 수 있도록 원본 행 보관
    return obj;
  });
}

function readFileAsText(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function () { resolve(reader.result); };
    reader.onerror = reject;
    reader.readAsText(file, "utf-8");
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function () { resolve(reader.result); };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// 엑셀(.xlsx) 파일의 첫 시트를 CSV 파싱 결과와 동일한 형태(헤더 기반 객체 배열)로 변환한다.
async function parseXlsx(file) {
  var buf = await readFileAsArrayBuffer(file);
  var wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  var ws = wb.worksheets[0];
  var rows = [];
  ws.eachRow(function (row) {
    var values = row.values.slice(1); // ExcelJS row.values는 1-based, index 0은 비움
    rows.push(values.map(function (v) {
      if (v === null || v === undefined) return "";
      if (v instanceof Date) return v.toISOString();
      if (typeof v === "object" && v.text !== undefined) return v.text; // richtext
      if (typeof v === "object" && v.result !== undefined) return v.result; // formula
      return String(v);
    }));
  });
  if (rows.length === 0) return [];
  var headers = rows[0];
  return rows.slice(1).map(function (r) {
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = r[idx] !== undefined ? r[idx] : ""; });
    obj.__cells = r; // 연차 시간(F/G열)을 헤더명과 무관하게 열 위치로 읽을 수 있도록 원본 행 보관
    return obj;
  });
}

function isExcelFile(file) {
  return /\.xlsx$/i.test(file.name);
}

// 파일 확장자에 따라 CSV 또는 XLSX를 읽어 동일한 형태(헤더 기반 객체 배열)로 반환한다.
async function parseSpreadsheetFile(file) {
  if (isExcelFile(file)) return parseXlsx(file);
  return parseCsv(await readFileAsText(file));
}

function toFloat(v) {
  if (v === undefined || v === null || v === "") return 0;
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
