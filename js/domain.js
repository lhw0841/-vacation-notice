// 연차 데이터 병합/그룹핑/기간 계산 등 업무 로직 (generate_vacation_notice.py와 1:1 대응)

import { toFloat } from "./csv.js";

export function loadEmployees(less1Granted, less1Used, tenuredGranted, tenuredUsed) {
  var usedById = {};
  less1Used.forEach(function (r) { usedById[r["아이디"]] = r; });
  tenuredUsed.forEach(function (r) { usedById[r["아이디"]] = r; });

  var employees = [];
  less1Granted.forEach(function (r) {
    var u = usedById[r["아이디"]] || {};
    employees.push({
      id: r["아이디"], name: r["이름"], dept: r["소속"], title: r["직위"],
      hire: r["입사일자"],
      granted: toFloat(r["1년 미만 연차"]),
      used: toFloat(u["1년 미만 연차(사용일)"]),
    });
  });
  tenuredGranted.forEach(function (r) {
    var u = usedById[r["아이디"]] || {};
    employees.push({
      id: r["아이디"], name: r["이름"], dept: r["소속"], title: r["직위"],
      hire: r["입사일자"],
      granted: toFloat(r["근속 연차"]),
      used: toFloat(u["근속 연차(사용일)"]),
    });
  });
  return employees;
}

export function groupByDept(employees) {
  var groups = {};
  employees.forEach(function (e) {
    if (!groups[e.dept]) groups[e.dept] = [];
    groups[e.dept].push(e);
  });
  Object.keys(groups).forEach(function (k) {
    groups[k].sort(function (a, b) { return a.name.localeCompare(b.name, "ko"); });
  });
  var sortedKeys = Object.keys(groups).sort(function (a, b) { return a.localeCompare(b, "ko"); });
  return sortedKeys.map(function (k) { return { dept: k, members: groups[k] }; });
}

function parseHireDate(s) {
  var parts = s.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function safeReplaceYear(d, year) {
  var month = d.getMonth(), day = d.getDate();
  var nd = new Date(year, month, day);
  if (nd.getMonth() !== month) nd = new Date(year, month, 28); // 2/29 -> 비윤년 예외
  return nd;
}

function fmtDate(d) {
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, "0");
  var day = String(d.getDate()).padStart(2, "0");
  return y + "." + m + "." + day;
}

// 통보일 기준 가장 최근 입사기념일 ~ 그다음 기념일 전날
export function periodStr(hireStr, noticeDate) {
  var hire = parseHireDate(hireStr);
  var annivThisYear = safeReplaceYear(hire, noticeDate.getFullYear());
  var start;
  if (hire > noticeDate) start = hire;
  else if (annivThisYear <= noticeDate) start = annivThisYear;
  else start = safeReplaceYear(hire, noticeDate.getFullYear() - 1);
  var end = safeReplaceYear(start, start.getFullYear() + 1);
  end.setDate(end.getDate() - 1);
  return fmtDate(start) + "-" + fmtDate(end);
}

// 통보일 다음달부터 1년(12개월): 예) 2026년 7월 통보 -> 2026.08 ~ 2027.07
export function usagePeriodMonths(noticeDate) {
  var y = noticeDate.getFullYear();
  var m = noticeDate.getMonth() + 1; // 1-12
  var result = [];
  for (var i = 0; i < 12; i++) {
    m += 1;
    if (m > 12) { m = 1; y += 1; }
    result.push({ year: y, month: m });
  }
  return result;
}

export function safeSheetName(dept) {
  return dept.replace(/[\\/*?:\[\]]/g, "_").slice(0, 31);
}

export function safeFileName(dept) {
  return dept.replace(/[\\/*?:"<>|]/g, "_");
}
