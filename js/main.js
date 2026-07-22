// 화면(UI) 연결부: 파일 선택 상태 관리, 버튼 클릭 처리, 결과 렌더링.
// JSZip은 vendor/jszip.min.js가 만든 전역(window.JSZip)을 사용한다.

import { parseCsv, readFileAsText } from "./csv.js";
import { loadEmployees, groupByDept, safeFileName } from "./domain.js";
import { buildWorkbook } from "./excel.js";

var fileState = {
  less1granted: null,
  less1used: null,
  tenuredgranted: null,
  tenuredused: null,
};
var generatedFiles = []; // {dept, count, filename, blob}

var fileInputMap = {
  less1granted: "f-less1granted",
  less1used: "f-less1used",
  tenuredgranted: "f-tenuredgranted",
  tenuredused: "f-tenuredused",
};

Object.keys(fileInputMap).forEach(function (key) {
  var input = document.getElementById(fileInputMap[key]);
  var box = input.closest(".filebox");
  var nameEl = box.querySelector("[data-name]");
  input.addEventListener("change", function () {
    var file = input.files[0];
    fileState[key] = file || null;
    if (file) {
      nameEl.textContent = file.name;
      nameEl.classList.remove("empty");
      box.classList.add("has-file");
    } else {
      nameEl.textContent = "파일 선택 안 됨";
      nameEl.classList.add("empty");
      box.classList.remove("has-file");
    }
  });
});

var noticeDateInput = document.getElementById("notice-date");
(function initDate() {
  var today = new Date();
  var iso = today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" + String(today.getDate()).padStart(2, "0");
  noticeDateInput.value = iso;
})();

var statusEl = document.getElementById("status");
var runBtn = document.getElementById("run-btn");
var resultsEl = document.getElementById("results");
var resultRowsEl = document.getElementById("result-rows");
var zipBtn = document.getElementById("zip-btn");

function setStatus(msg, isError, busy) {
  statusEl.innerHTML = "";
  statusEl.classList.toggle("error", !!isError);
  if (busy) {
    var sp = document.createElement("div");
    sp.className = "spinner";
    statusEl.appendChild(sp);
  }
  var span = document.createElement("span");
  span.textContent = msg;
  statusEl.appendChild(span);
}

function triggerDownload(blob, filename) {
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
}

runBtn.addEventListener("click", async function () {
  var missing = Object.keys(fileState).filter(function (k) { return !fileState[k]; });
  if (missing.length) {
    setStatus("CSV 파일 4개를 모두 선택해 주세요.", true, false);
    return;
  }
  var noticeDateVal = noticeDateInput.value;
  if (!noticeDateVal) {
    setStatus("회사통보일을 선택해 주세요.", true, false);
    return;
  }

  runBtn.disabled = true;
  resultsEl.classList.remove("show");
  generatedFiles = [];

  try {
    setStatus("CSV 파일을 읽는 중...", false, true);
    var texts = await Promise.all([
      readFileAsText(fileState.less1granted),
      readFileAsText(fileState.less1used),
      readFileAsText(fileState.tenuredgranted),
      readFileAsText(fileState.tenuredused),
    ]);
    var less1Granted = parseCsv(texts[0]);
    var less1Used = parseCsv(texts[1]);
    var tenuredGranted = parseCsv(texts[2]);
    var tenuredUsed = parseCsv(texts[3]);

    var employees = loadEmployees(less1Granted, less1Used, tenuredGranted, tenuredUsed);
    if (!employees.length) throw new Error("CSV에서 직원 데이터를 찾지 못했습니다. 파일 형식을 확인해 주세요.");

    var groups = groupByDept(employees);
    var dateParts = noticeDateVal.split("-").map(Number);
    var noticeDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);

    resultRowsEl.innerHTML = "";
    for (var i = 0; i < groups.length; i++) {
      var g = groups[i];
      setStatus("생성 중... (" + (i + 1) + "/" + groups.length + ") " + g.dept, false, true);
      var wb = await buildWorkbook(g.dept, g.members, noticeDate);
      var buf = await wb.xlsx.writeBuffer();
      var blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      var filename = "미사용 연차유급휴가 사용시기 지정통보서_" + safeFileName(g.dept) + ".xlsx";
      generatedFiles.push({ dept: g.dept, count: g.members.length, filename: filename, blob: blob });

      var tr = document.createElement("tr");
      var td1 = document.createElement("td"); td1.textContent = g.dept;
      var td2 = document.createElement("td"); td2.textContent = g.members.length + "명"; td2.className = "num";
      var td3 = document.createElement("td"); td3.textContent = "완료";
      var td4 = document.createElement("td");
      var dlBtn = document.createElement("button");
      dlBtn.className = "btn btn-ghost btn-sm";
      dlBtn.textContent = "다운로드";
      dlBtn.addEventListener("click", (function (f) {
        return function () { triggerDownload(f.blob, f.filename); };
      })(generatedFiles[generatedFiles.length - 1]));
      td4.appendChild(dlBtn);
      tr.appendChild(td1); tr.appendChild(td2); tr.appendChild(td3); tr.appendChild(td4);
      resultRowsEl.appendChild(tr);
    }

    document.getElementById("stat-people").textContent = employees.length;
    document.getElementById("stat-depts").textContent = groups.length;
    document.getElementById("stat-date").textContent = noticeDateVal.replace(/-/g, ".");
    resultsEl.classList.add("show");

    setStatus("완료: " + employees.length + "명, " + groups.length + "개 부서 파일 생성 완료.", false, false);
  } catch (err) {
    console.error(err);
    setStatus("오류: " + err.message, true, false);
  } finally {
    runBtn.disabled = false;
  }
});

zipBtn.addEventListener("click", async function () {
  if (!generatedFiles.length) return;
  zipBtn.disabled = true;
  var origText = zipBtn.textContent;
  zipBtn.textContent = "압축 중...";
  try {
    var zip = new JSZip();
    generatedFiles.forEach(function (f) { zip.file(f.filename, f.blob); });
    var content = await zip.generateAsync({ type: "blob" });
    triggerDownload(content, "연차유급휴가_사용시기_지정통보서_부서별.zip");
  } finally {
    zipBtn.disabled = false;
    zipBtn.textContent = origText;
  }
});
