// ExcelJS로 부서별 통보서 워크북을 생성한다 (python 버전의 build_workbook과 1:1 대응)
// ExcelJS는 vendor/exceljs.min.js가 만든 전역(window.ExcelJS)을 사용한다.

var FONT_NAME = "맑은 고딕";
var THIN = { style: "thin" };
var BOX = { top: THIN, bottom: THIN, left: THIN, right: THIN };
var FILL_HILITE = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF6DFC8" } };
var FILL_GREEN = { type: "pattern", pattern: "solid", fgColor: { argb: "FFECFFEB" } };

function styleCell(cell, opts) {
  opts = opts || {};
  cell.font = opts.font || { name: FONT_NAME, size: 11 };
  cell.alignment = opts.align || { horizontal: "center", vertical: "middle" };
  cell.border = opts.border !== undefined ? opts.border : BOX;
  if (opts.fill) cell.fill = opts.fill;
}

function styleMergedRange(ws, r1, c1, r2, c2, opts) {
  for (var r = r1; r <= r2; r++) {
    for (var c = c1; c <= c2; c++) {
      styleCell(ws.getRow(r).getCell(c), opts);
    }
  }
}

var BOX_DOUBLE = { top: { style: "double" }, bottom: { style: "double" }, left: { style: "double" }, right: { style: "double" } };

async function buildWorkbook(dept, members, noticeDate) {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet(safeSheetName(dept));

  ws.getColumn(1).width = 17;
  ws.getColumn(2).width = 18;
  for (var c = 3; c <= 8; c++) ws.getColumn(c).width = 20;

  // 회사통보일
  ws.mergeCells(1, 1, 1, 2);
  ws.getCell(1, 1).value = "회사통보일 : " + noticeDate.getFullYear() + "년  " +
    String(noticeDate.getMonth() + 1).padStart(2, "0") + "월   " +
    String(noticeDate.getDate()).padStart(2, "0") + "일";
  styleMergedRange(ws, 1, 1, 1, 2, { font: { name: FONT_NAME, size: 12, bold: true }, border: {} });
  ws.getRow(1).height = 20;

  // 제목
  ws.mergeCells(3, 1, 5, 8);
  ws.getCell(3, 1).value = "<미사용 연차유급휴가 사용 시기 통보>";
  styleMergedRange(ws, 3, 1, 5, 8, { font: { name: FONT_NAME, size: 28 }, border: {} });
  ws.getRow(3).height = 17.25;

  // 안내문
  ws.mergeCells(8, 1, 14, 8);
  ws.getCell(8, 1).value =
    "기 통보받은 사용시기 지정통보서 내용을 바탕으로 연차유급휴가 사용 독려를 위한 통보서 이오니 꼭 회신 바랍니다.\n\n" +
    " 기 통보한 연차유급휴가 사용시기에 대한 수정이 필요한 경우 필히 서면 조치가 필요한 사항이오니 연락 후 조치 하시길 바랍니다.\n" +
    "  - 해당 절차는 근로기준법 제 61조 연차 유급휴가의 사용촉진 규정에 의거 통보되는 사항입니다.";
  styleMergedRange(ws, 8, 1, 14, 8, {
    font: { name: FONT_NAME, size: 16, bold: true },
    align: { horizontal: "center", vertical: "middle", wrapText: true },
    border: BOX_DOUBLE,
    fill: FILL_GREEN,
  });
  ws.getRow(8).height = 18.95;

  // 1. 통지받은 내역
  ws.getCell(18, 1).value = "1. 통지받은 내역";
  styleCell(ws.getCell(18, 1), { font: { name: FONT_NAME, size: 14, bold: true }, align: { vertical: "middle" }, border: {} });

  var headerRow = 20;
  ws.mergeCells(headerRow, 3, headerRow, 4);
  styleMergedRange(ws, headerRow, 3, headerRow, 4, { font: { name: FONT_NAME, size: 11, bold: true } });
  var headers = ["부서", "성명", "연차휴가 발생(사용) 대상 기간", null, "발생 연차(A)", "사용연차(B)", "미사용 연차(A-B)", "확인(사인)"];
  headers.forEach(function (text, idx) {
    if (text === null) return;
    var col = idx + 1;
    var cell = ws.getCell(headerRow, col);
    cell.value = text;
    var fill = (col === 5 || col === 7) ? FILL_HILITE : undefined;
    styleCell(cell, { font: { name: FONT_NAME, size: 11, bold: true }, fill: fill });
  });
  ws.getRow(headerRow).height = 35;

  var dataStart = headerRow + 1;
  members.forEach(function (emp, i) {
    var r = dataStart + i;
    ws.mergeCells(r, 3, r, 4);
    styleMergedRange(ws, r, 3, r, 4, { font: { name: FONT_NAME, size: 11, bold: false } });

    styleCell(ws.getCell(r, 1), { font: { name: FONT_NAME, size: 11, bold: true } });
    ws.getCell(r, 1).value = dept;

    styleCell(ws.getCell(r, 2), { font: { name: FONT_NAME, size: 11, bold: false } });
    ws.getCell(r, 2).value = emp.name;

    ws.getCell(r, 3).value = periodStr(emp.hire, noticeDate);

    styleCell(ws.getCell(r, 5), { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });
    ws.getCell(r, 5).value = emp.granted;

    styleCell(ws.getCell(r, 6), { font: { name: FONT_NAME, size: 11, bold: true } });
    ws.getCell(r, 6).value = emp.used;

    styleCell(ws.getCell(r, 7), { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });
    ws.getCell(r, 7).value = { formula: "E" + r + "-F" + r };

    styleCell(ws.getCell(r, 8), { font: { name: FONT_NAME, size: 11, bold: true } });

    ws.getRow(r).height = 35;
  });

  // 2. 개인별 사용시기 통보 — 미사용 연차(A-B) 개수만큼 무작위 영업일을 뽑아 "M월 D일 / M월 D일 ..." 한 줄로 나열
  var section2Row = dataStart + members.length + 3;
  ws.getCell(section2Row, 1).value = "2. 미사용 연차유급휴가 개인별 사용시기 통보";
  styleCell(ws.getCell(section2Row, 1), { font: { name: FONT_NAME, size: 14, bold: true }, align: { vertical: "middle" }, border: {} });

  var table2Header = section2Row + 2;
  ws.mergeCells(table2Header, 2, table2Header, 8);
  ws.getCell(table2Header, 1).value = "성명";
  styleCell(ws.getCell(table2Header, 1), { font: { name: FONT_NAME, size: 11, bold: true } });
  ws.getCell(table2Header, 2).value = "사용시기 지정일";
  styleMergedRange(ws, table2Header, 2, table2Header, 8, { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });
  ws.getRow(table2Header).height = 25;

  // 넓힌 열(2~8열, 총 138 width) 기준 한 줄에 대략 들어가는 날짜 개수 — 잘림 방지를 위해 줄바꿈 + 여유있는 행 높이 적용
  var ENTRIES_PER_LINE = 7;
  var row = table2Header + 1;
  members.forEach(function (emp) {
    var count = Math.max(0, Math.round(emp.granted - emp.used));
    var lines = Math.max(1, Math.ceil(count / ENTRIES_PER_LINE));

    ws.mergeCells(row, 2, row, 8);
    styleCell(ws.getCell(row, 1), { font: { name: FONT_NAME, size: 11 } });
    ws.getCell(row, 1).value = emp.name;
    styleMergedRange(ws, row, 2, row, 8, {
      font: { name: FONT_NAME, size: 11 },
      align: { horizontal: "left", vertical: "middle", wrapText: true },
    });
    ws.getCell(row, 2).value = usageDateListStr(noticeDate, count);
    ws.getRow(row).height = Math.max(25, lines * 18);
    row++;
  });

  return wb;
}

// 1차 통보 서식: 실제 사내 양식(1. 미사용 연차유급휴가 사용 시기 통보_○○부.xlsx)과 동일한 레이아웃.
// 2번 표는 12개월 표가 아니라, 직원별 잔여연차(미사용 연차) 개수만큼 주말을 제외한 평일을
// 통보일 다음날부터 1년 내에서 무작위로 뽑아 "M월 D일 / M월 D일 ..." 한 줄로 나열한다.
async function buildWorkbookV1(dept, members, noticeDate) {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet(safeSheetName(dept));

  ws.getColumn(2).width = 18;
  for (var c = 3; c <= 9; c++) ws.getColumn(c).width = 20;

  // 회사통보일
  ws.mergeCells(1, 2, 1, 3);
  ws.getCell(1, 2).value = "회사통보일 : " + noticeDate.getFullYear() + "년  " +
    String(noticeDate.getMonth() + 1).padStart(2, "0") + "월   " +
    String(noticeDate.getDate()).padStart(2, "0") + "일";
  styleMergedRange(ws, 1, 2, 1, 3, { font: { name: FONT_NAME, size: 12, bold: true }, border: {} });

  // 제목
  ws.mergeCells(3, 2, 5, 9);
  ws.getCell(3, 2).value = "<미사용 연차유급휴가 사용 시기 통보>";
  styleMergedRange(ws, 3, 2, 5, 9, { font: { name: FONT_NAME, size: 28 }, border: {} });
  ws.getRow(3).height = 17.25;

  // 안내문
  ws.mergeCells(8, 2, 14, 9);
  ws.getCell(8, 2).value =
    "기 통보받은 사용시기 지정통보서 내용을 바탕으로 연차유급휴가 사용 독려를 위한 통보서 이오니 꼭 회신 바랍니다.\n\n" +
    " 기 통보한 연차유급휴가 사용시기에 대한 수정이 필요한 경우 필히 서면 조치가 필요한 사항이오니 연락 후 조치 하시길 바랍니다.\n" +
    "  - 해당 절차는 근로기준법 제 61조 연차 유급휴가의 사용촉진 규정에 의거 통보되는 사항입니다.";
  styleMergedRange(ws, 8, 2, 14, 9, {
    font: { name: FONT_NAME, size: 16, bold: true },
    align: { horizontal: "center", vertical: "middle", wrapText: true },
    border: BOX_DOUBLE,
    fill: FILL_GREEN,
  });
  ws.getRow(8).height = 18.95;

  // 1. 통지받은 내역
  ws.getCell(18, 2).value = "1. 통지받은 내역";
  styleCell(ws.getCell(18, 2), { font: { name: FONT_NAME, size: 14, bold: true }, align: { vertical: "middle" }, border: {} });

  var headerRow = 20;
  ws.mergeCells(headerRow, 4, headerRow, 5);
  var headers = ["부서", "성명", "연차휴가 발행(사용) 대상 기간", null, "발생 연차(A)", "사용연차(B)", "미사용 연차(A-B)", "확인(사인)"];
  headers.forEach(function (text, idx) {
    if (text === null) return;
    var col = idx + 2;
    var cell = ws.getCell(headerRow, col);
    cell.value = text;
    var fill = (col === 6 || col === 8 || col === 9) ? FILL_HILITE : undefined;
    styleCell(cell, { font: { name: FONT_NAME, size: 11, bold: true }, fill: fill });
  });
  styleMergedRange(ws, headerRow, 4, headerRow, 5, { font: { name: FONT_NAME, size: 11, bold: true } });
  ws.getRow(headerRow).height = 35.1;

  var dataStart = headerRow + 1;
  members.forEach(function (emp, i) {
    var r = dataStart + i;
    ws.mergeCells(r, 4, r, 5);
    styleMergedRange(ws, r, 4, r, 5, { font: { name: FONT_NAME, size: 11, bold: false } });

    styleCell(ws.getCell(r, 2), { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });
    ws.getCell(r, 2).value = dept;

    styleCell(ws.getCell(r, 3), { font: { name: FONT_NAME, size: 11, bold: false } });
    ws.getCell(r, 3).value = emp.name;

    ws.getCell(r, 4).value = periodStr(emp.hire, noticeDate);

    styleCell(ws.getCell(r, 6), { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });
    ws.getCell(r, 6).value = emp.granted;

    styleCell(ws.getCell(r, 7), { font: { name: FONT_NAME, size: 11, bold: true } });
    ws.getCell(r, 7).value = emp.used;

    styleCell(ws.getCell(r, 8), { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });
    ws.getCell(r, 8).value = { formula: "F" + r + "-G" + r };

    styleCell(ws.getCell(r, 9), { font: { name: FONT_NAME, size: 11, bold: true }, fill: FILL_HILITE });

    ws.getRow(r).height = 35.1;
  });

  // 2. 미사용 연차유급휴가 사용시기 통보 (직접 작성용, 빈칸)
  var section2Row = dataStart + members.length + 3;
  ws.getCell(section2Row, 2).value = "2. 미사용 연차유급휴가 사용시기 통보";
  styleCell(ws.getCell(section2Row, 2), { font: { name: FONT_NAME, size: 14, bold: true }, align: { vertical: "middle" }, border: {} });

  var table2Header = section2Row + 2;
  ws.mergeCells(table2Header, 3, table2Header, 9);
  ws.getCell(table2Header, 2).value = "성명";
  styleCell(ws.getCell(table2Header, 2), { font: { name: FONT_NAME, size: 14, bold: true } });
  ws.getCell(table2Header, 3).value = "기 통보 사용시기 지정일";
  styleMergedRange(ws, table2Header, 3, table2Header, 9, { font: { name: FONT_NAME, size: 14, bold: true }, fill: FILL_HILITE });
  ws.getRow(table2Header).height = 24.95;

  var row = table2Header + 1;
  members.forEach(function (emp) {
    ws.mergeCells(row, 3, row, 9);
    styleCell(ws.getCell(row, 2), { font: { name: FONT_NAME, size: 11 } });
    ws.getCell(row, 2).value = emp.name;
    styleMergedRange(ws, row, 3, row, 9, { font: { name: FONT_NAME, size: 11 } });
    ws.getRow(row).height = 24.95;
    row++;
  });

  return wb;
}
