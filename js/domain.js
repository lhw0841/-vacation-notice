// 연차 데이터 병합/그룹핑/기간 계산 등 업무 로직 (generate_vacation_notice.py와 1:1 대응)

function loadEmployees(less1Granted, less1Used, tenuredGranted, tenuredUsed) {
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

function groupByDept(employees) {
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
function periodStr(hireStr, noticeDate) {
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
function usagePeriodMonths(noticeDate) {
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

function isWeekday(d) {
  var wd = d.getDay();
  return wd !== 0 && wd !== 6;
}

// 2026~2029년 대한민국 법정공휴일(대체공휴일 포함, 설/추석 등 음력 명절 포함) — "YYYY-MM-DD" 문자열 목록.
// 음력 기반 명절(설날/추석/부처님오신날)은 해마다 날짜가 바뀌므로 이 표는 매년 직접 갱신이 필요하다.
// 여기 없는 해는 공휴일 제외 없이 주말만 걸러진다.
var KOREAN_HOLIDAYS = {
  "2026-01-01": 1, // 신정
  "2026-02-16": 1, "2026-02-17": 1, "2026-02-18": 1, // 설날 연휴
  "2026-03-01": 1, "2026-03-02": 1, // 삼일절(대체)
  "2026-05-01": 1, // 근로자의 날
  "2026-05-05": 1, // 어린이날
  "2026-05-24": 1, "2026-05-25": 1, // 부처님오신날(대체)
  "2026-06-06": 1, // 현충일
  "2026-07-17": 1, // 제헌절
  "2026-08-15": 1, "2026-08-17": 1, // 광복절(대체)
  "2026-09-24": 1, "2026-09-25": 1, "2026-09-26": 1, // 추석 연휴
  "2026-10-03": 1, // 개천절
  "2026-10-09": 1, // 한글날
  "2026-12-25": 1, // 크리스마스

  "2027-01-01": 1, // 신정
  "2027-02-06": 1, "2027-02-07": 1, "2027-02-08": 1, "2027-02-09": 1, // 설날 연휴(대체)
  "2027-03-01": 1, // 삼일절
  "2027-05-01": 1, // 근로자의 날
  "2027-05-05": 1, // 어린이날
  "2027-05-13": 1, // 부처님오신날
  "2027-06-06": 1, // 현충일
  "2027-07-17": 1, // 제헌절
  "2027-08-15": 1, "2027-08-16": 1, // 광복절(대체)
  "2027-09-14": 1, "2027-09-15": 1, "2027-09-16": 1, // 추석 연휴
  "2027-10-03": 1, "2027-10-04": 1, // 개천절(대체)
  "2027-10-09": 1, "2027-10-11": 1, // 한글날(대체)
  "2027-12-25": 1, "2027-12-27": 1, // 크리스마스(대체)

  "2028-01-01": 1, // 신정
  "2028-01-26": 1, "2028-01-27": 1, "2028-01-28": 1, // 설날 연휴
  "2028-03-01": 1, // 삼일절
  "2028-05-01": 1, // 근로자의 날
  "2028-05-02": 1, // 부처님오신날
  "2028-05-05": 1, // 어린이날
  "2028-06-06": 1, // 현충일
  "2028-07-17": 1, // 제헌절
  "2028-08-15": 1, // 광복절
  "2028-10-02": 1, "2028-10-03": 1, "2028-10-04": 1, // 추석 연휴(10/3 개천절과 겹침)
  "2028-10-05": 1, // 추석·개천절 겹침 대체공휴일
  "2028-10-09": 1, // 한글날
  "2028-12-25": 1, // 크리스마스

  "2029-01-01": 1, // 신정
  "2029-02-12": 1, "2029-02-13": 1, "2029-02-14": 1, // 설날 연휴
  "2029-03-01": 1, // 삼일절
  "2029-05-01": 1, // 근로자의 날
  "2029-05-05": 1, "2029-05-07": 1, // 어린이날(토, 대체 월)
  "2029-05-20": 1, "2029-05-21": 1, // 부처님오신날(일, 대체 월)
  "2029-06-06": 1, // 현충일
  "2029-07-17": 1, // 제헌절
  "2029-08-15": 1, // 광복절
  "2029-09-21": 1, "2029-09-22": 1, "2029-09-23": 1, "2029-09-24": 1, // 추석 연휴(토 겹침, 대체 월)
  "2029-10-03": 1, // 개천절
  "2029-10-09": 1, // 한글날
  "2029-12-25": 1, // 크리스마스
};

function isoDate(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

function isHoliday(d) {
  return !!KOREAN_HOLIDAYS[isoDate(d)];
}

function isBusinessDay(d) {
  return isWeekday(d) && !isHoliday(d);
}

// 통보일 다음날부터 1년간의 영업일(주말+공휴일 제외) 중 count개를 무작위로 뽑아 오름차순으로 반환
function pickRandomBusinessDates(noticeDate, count) {
  count = Math.max(0, Math.round(count));
  var start = new Date(noticeDate.getFullYear(), noticeDate.getMonth(), noticeDate.getDate() + 1);
  var end = new Date(noticeDate.getFullYear() + 1, noticeDate.getMonth(), noticeDate.getDate() - 1);
  var pool = [];
  for (var d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (isBusinessDay(d)) pool.push(new Date(d));
  }
  for (var i = pool.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = pool[i]; pool[i] = pool[j]; pool[j] = tmp;
  }
  var picked = pool.slice(0, Math.min(count, pool.length));
  picked.sort(function (a, b) { return a - b; });
  return picked;
}

// "10월 28일 / 10월 29일 / 2026년 1월 5일 / ..." 형식으로 잔여연차 개수만큼 지정일 목록 문자열 생성.
// 연도가 바뀌는 지점에서만 "YYYY년"을 붙여 표시한다(대상기간이 해를 넘길 수 있어 혼동 방지).
function usageDateListStr(noticeDate, remainCount) {
  var lastYear = null;
  return pickRandomBusinessDates(noticeDate, remainCount)
    .map(function (d) {
      var y = d.getFullYear();
      var label = (d.getMonth() + 1) + "월 " + d.getDate() + "일";
      if (y !== lastYear) {
        label = y + "년 " + label;
        lastYear = y;
      }
      return label;
    })
    .join(" / ");
}

function safeSheetName(dept) {
  return dept.replace(/[\\/*?:\[\]]/g, "_").slice(0, 31);
}

function safeFileName(dept) {
  return dept.replace(/[\\/*?:"<>|]/g, "_");
}
