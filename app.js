/* global XLSX, pdfjsLib */

const RULES = {
  salarySheetKeywords: ["교직원", "보수", "일람", "기준"],
  retirementSheetKeywords: ["퇴직", "적립", "충당"],
  teacherKeywords: ["원장", "원감", "교사", "담임", "방과후"],
  regularTeacherKeywords: ["원장", "원감", "교사", "담임"],
  afterSchoolKeywords: ["방과후"],
  staffKeywords: ["직원", "사무", "조리", "보조", "기사", "환경", "영양"],
  allowanceHeaders: ["수당", "급식", "명절", "상여", "휴가", "성과", "자가", "처우", "담임", "직책"],
  basePayHeaders: ["본봉", "급여", "기본급"],
  personnelExpenseKeywords: ["인건비", "급여", "수당", "법정부담금", "4대보험", "사학연금", "퇴직", "적립금", "상여금"],
  allowedPersonnelBuckets: ["교원인건비", "직원인건비", "그밖의인건비"],
  pdfBudgetItems: {
    teacherPay: ["교원급여"],
    afterSchoolTeacherPay: ["방과후교원급여", "방과후급여"],
    teacherAllowance: ["교원수당"],
    teacherMealAllowance: ["교원정액급식비", "정액급식비"],
    staffPay: ["직원급여"],
    staffAllowance: ["직원수당"],
    driverPay: ["차량기사급여"],
    retirement: ["퇴직", "퇴직금", "퇴직적립", "퇴직급여충당"]
  }
};

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const $ = (id) => document.getElementById(id);
const fmt = (n) => Number.isFinite(n) ? n.toLocaleString("ko-KR") : "-";
const norm = (s) => String(s ?? "").replace(/\s+/g, "").trim();
const toNumber = (v) => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const cleaned = String(v ?? "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
};

$("sampleRulesBtn").addEventListener("click", () => $("rulesDialog").showModal());
$("closeRulesBtn").addEventListener("click", () => $("rulesDialog").close());
$("runBtn").addEventListener("click", runReview);

async function runReview() {
  const excelFile = $("excelFile").files[0];
  const pdfFile = $("pdfFile").files[0];
  if (!excelFile || !pdfFile) return setStatus("엑셀 파일과 PDF 파일을 모두 업로드해주세요.", true);

  setStatus("파일을 읽고 있습니다...");
  try {
    const [excel, pdf] = await Promise.all([parseExcel(excelFile), parsePdf(pdfFile)]);
    setStatus("검토 규칙을 적용하고 있습니다...");
    const report = analyze(excel, pdf);
    renderReport(report);
    setStatus("검토가 완료되었습니다.");
  } catch (err) {
    console.error(err);
    setStatus(`오류가 발생했습니다: ${err.message}`, true);
  }
}

function setStatus(message, isError = false) {
  const el = $("status");
  el.hidden = false;
  el.textContent = message;
  el.style.background = isError ? "#fef3f2" : "#eff8ff";
  el.style.color = isError ? "#b42318" : "#1849a9";
}

async function parseExcel(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array", cellDates: false, cellFormula: true });
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    return { name, matrix, text: matrix.flat().map(String).join(" ") };
  });
  const salarySheet = pickSheet(sheets, RULES.salarySheetKeywords) || sheets[0];
  const retirementSheets = sheets.filter(s => hasAny(s.name + " " + s.text, RULES.retirementSheetKeywords));
  const salaryRows = inferSalaryRows(salarySheet.matrix);
  return { sheets, salarySheet, salaryRows, retirementSheets };
}

function pickSheet(sheets, keywords) {
  return sheets
    .map(s => ({ sheet: s, score: keywords.reduce((a, k) => a + (norm(s.name).includes(norm(k)) ? 3 : 0) + (norm(s.text).includes(norm(k)) ? 1 : 0), 0) }))
    .sort((a, b) => b.score - a.score)[0]?.score > 0
    ? sheets.map(s => ({ sheet: s, score: keywords.reduce((a, k) => a + (norm(s.name).includes(norm(k)) ? 3 : 0) + (norm(s.text).includes(norm(k)) ? 1 : 0), 0) })).sort((a, b) => b.score - a.score)[0].sheet
    : null;
}

function inferSalaryRows(matrix) {
  const headerRowIndex = matrix.findIndex(row => row.some(c => RULES.basePayHeaders.concat(RULES.allowanceHeaders).some(k => norm(c).includes(norm(k)))));
  const headers = headerRowIndex >= 0 ? matrix[headerRowIndex].map((h, i) => ({ name: String(h || `열${i + 1}`), index: i })) : [];
  const amountColumns = headers.filter(h => RULES.basePayHeaders.concat(RULES.allowanceHeaders).some(k => norm(h.name).includes(norm(k))));
  const rows = [];
  for (let r = Math.max(0, headerRowIndex + 1); r < matrix.length; r++) {
    const row = matrix[r];
    const rowText = row.map(String).join(" ");
    const hasAmount = amountColumns.some(c => toNumber(row[c.index]) > 0);
    if (!hasAmount) continue;
    const amounts = {};
    amountColumns.forEach(c => { amounts[c.name] = toNumber(row[c.index]); });
    rows.push({ rowNumber: r + 1, rowText, amounts });
  }
  return rows;
}

async function parsePdf(file) {
  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const items = content.items.map(item => ({ str: item.str, x: item.transform[4], y: Math.round(item.transform[5]) }));
    const linesMap = new Map();
    items.forEach(item => {
      const key = item.y;
      if (!linesMap.has(key)) linesMap.set(key, []);
      linesMap.get(key).push(item);
    });
    const lines = [...linesMap.entries()]
      .sort((a, b) => b[0] - a[0])
      .map(([, arr]) => arr.sort((a, b) => a.x - b.x).map(x => x.str).join(" ").trim())
      .filter(Boolean);
    pages.push({ page: i, lines, text: lines.join("\n") });
  }
  const fullText = pages.map(p => p.text).join("\n");
  return { pages, fullText, compactText: norm(fullText) };
}

function analyze(excel, pdf) {
  const results = [];
  const debug = [];
  const salaryRows = excel.salaryRows;

  debug.push(`선택된 보수 시트: ${excel.salarySheet?.name || "확인 불가"}`);
  debug.push(`보수 행 후보: ${salaryRows.length}건`);
  debug.push(`퇴직 관련 시트 후보: ${excel.retirementSheets.map(s => s.name).join(", ") || "없음"}`);

  const groups = buildExcelGroups(salaryRows);
  const pdfItems = extractPdfItems(pdf.fullText);
  debug.push(`PDF 산출기초 후보: ${pdfItems.length}건`);

  comparePay(results, "교원급여", groups.teacherPay, findPdfAmount(pdf, RULES.pdfBudgetItems.teacherPay), findPdfPeopleNear(pdf, RULES.pdfBudgetItems.teacherPay), true);
  comparePay(results, "방과후교원급여", groups.afterSchoolTeacherPay, findPdfAmount(pdf, RULES.pdfBudgetItems.afterSchoolTeacherPay), findPdfPeopleNear(pdf, RULES.pdfBudgetItems.afterSchoolTeacherPay), true);
  comparePay(results, "직원급여", groups.staffPay, findPdfAmount(pdf, RULES.pdfBudgetItems.staffPay), findPdfPeopleNear(pdf, RULES.pdfBudgetItems.staffPay), false);

  checkAllowances(results, groups.teacherAllowances, pdf, "교원수당");
  checkAllowances(results, groups.staffAllowances, pdf, "직원수당");
  checkRetirement(results, excel, pdf);
  checkMisclassifiedPersonnel(results, pdf);

  return { results, debug: debug.concat(buildDebug(groups, pdfItems)) };
}

function buildExcelGroups(rows) {
  const sumByHeaders = (filterFn, headerFn) => {
    let amount = 0, people = 0;
    rows.filter(filterFn).forEach(row => {
      const rowSum = Object.entries(row.amounts).reduce((a, [h, v]) => a + (headerFn(h) ? v : 0), 0);
      if (rowSum > 0) { amount += rowSum; people++; }
    });
    return { amount, people };
  };
  const isRegularTeacher = r => hasAny(r.rowText, RULES.regularTeacherKeywords) && !hasAny(r.rowText, ["방과후"]);
  const isAfterSchool = r => hasAny(r.rowText, RULES.afterSchoolKeywords);
  const isStaff = r => hasAny(r.rowText, RULES.staffKeywords) && !hasAny(r.rowText, RULES.teacherKeywords);
  const isBase = h => hasAny(h, RULES.basePayHeaders);
  const isAllowance = h => hasAny(h, RULES.allowanceHeaders) && !isBase(h);
  return {
    teacherPay: sumByHeaders(isRegularTeacher, isBase),
    afterSchoolTeacherPay: sumByHeaders(isAfterSchool, isBase),
    staffPay: sumByHeaders(isStaff, isBase),
    teacherAllowances: collectAllowances(rows.filter(isRegularTeacher), isAllowance),
    staffAllowances: collectAllowances(rows.filter(isStaff), isAllowance)
  };
}

function collectAllowances(rows, headerFn) {
  const map = new Map();
  rows.forEach(row => {
    Object.entries(row.amounts).forEach(([header, value]) => {
      if (value > 0 && headerFn(header)) map.set(header, (map.get(header) || 0) + value);
    });
  });
  return [...map.entries()].map(([name, amount]) => ({ name, amount }));
}

function extractPdfItems(text) {
  const lines = text.split(/\n+/).map(s => s.trim()).filter(Boolean);
  return lines.filter(line => /원\s*\*/.test(line) || /=\s*[0-9,]+/.test(line));
}

function findPdfAmount(pdf, aliases) {
  for (const alias of aliases) {
    const idx = pdf.compactText.indexOf(norm(alias));
    if (idx < 0) continue;
    const slice = pdf.compactText.slice(idx, idx + 260);
    const formulaMatch = slice.match(/=([0-9,]+)원?/);
    if (formulaMatch) return toNumber(formulaMatch[1]);
    const tableMatch = slice.match(new RegExp(`${norm(alias)}([0-9,]{1,12})`));
    if (tableMatch) return toNumber(tableMatch[1]) * 1000;
  }
  return 0;
}

function findPdfPeopleNear(pdf, aliases) {
  for (const alias of aliases) {
    const idx = pdf.compactText.indexOf(norm(alias));
    if (idx < 0) continue;
    const slice = pdf.compactText.slice(idx, idx + 220);
    const matches = [...slice.matchAll(/\*([0-9]+)명\*/g)].map(m => Number(m[1]));
    if (matches.length) return matches.reduce((a, b) => a + b, 0);
  }
  return 0;
}

function comparePay(results, label, excelGroup, pdfAmount, pdfPeople, checkPeople) {
  const amountDiff = excelGroup.amount - pdfAmount;
  const peopleDiff = excelGroup.people - pdfPeople;
  let status = "ok", message = "금액 일치";
  if (excelGroup.amount === 0 && pdfAmount === 0) {
    status = "warn"; message = "자동 추출 실패 또는 항목 없음";
  } else if (Math.abs(amountDiff) > 1000) {
    status = "bad"; message = `금액 차이 ${fmt(amountDiff)}원`;
  }
  if (checkPeople && pdfPeople > 0 && peopleDiff !== 0) {
    status = "bad";
    message += ` / 인원수 불일치: PDF ${pdfPeople}명, 엑셀 ${excelGroup.people}명`;
  }
  results.push({ category: "보수", item: label, excelAmount: excelGroup.amount, pdfAmount, excelPeople: excelGroup.people, pdfPeople, status, message });
}

function checkAllowances(results, allowances, pdf, bucketName) {
  const missing = [];
  const direct = [];
  allowances.forEach(a => {
    const amount = findPdfAmount(pdf, [a.name]);
    if (amount > 0 && Math.abs(amount - a.amount) <= 1000) direct.push({ ...a, pdfAmount: amount });
    else missing.push(a);
  });
  direct.forEach(a => results.push({ category: "수당", item: a.name, excelAmount: a.amount, pdfAmount: a.pdfAmount, status: "ok", message: "개별 편성 일치" }));

  const missingSum = missing.reduce((a, b) => a + b.amount, 0);
  if (missing.length === 0) return;
  const bucketAmount = findPdfAmount(pdf, [bucketName]);
  if (bucketAmount && Math.abs(bucketAmount - missingSum) <= 1000) {
    results.push({ category: "수당", item: bucketName, excelAmount: missingSum, pdfAmount: bucketAmount, status: "info", message: `통합편성(${missing.map(x => `${x.name} ${fmt(x.amount)}원`).join(" + ")} = ${fmt(missingSum)}원)` });
  } else {
    results.push({ category: "수당", item: bucketName, excelAmount: missingSum, pdfAmount: bucketAmount, status: "warn", message: `추가확인필요: PDF에서 개별 확인되지 않은 수당 ${missing.length}건의 합계와 ${bucketName} 금액이 일치하지 않습니다.` });
  }
}

function checkRetirement(results, excel, pdf) {
  const hasRetirementAmount = excel.retirementSheets.some(s => s.matrix.flat().some(c => toNumber(c) > 0));
  if (!hasRetirementAmount) return;
  const found = hasAny(pdf.fullText, RULES.pdfBudgetItems.retirement);
  results.push({ category: "퇴직금", item: "퇴직적립금 편성 여부", status: found ? "ok" : "bad", message: found ? "퇴직 관련 편성 항목 확인" : "퇴직금 적립금 미편성 의심", excelAmount: null, pdfAmount: null });
}

function checkMisclassifiedPersonnel(results, pdf) {
  const lines = pdf.fullText.split(/\n+/).map(s => s.trim()).filter(Boolean);
  let currentBucket = "";
  const suspicious = [];
  for (const line of lines) {
    if (/인건비|급식비|통학차량|교육활동|운영비|시설비|적립금/.test(line) && !/=/.test(line)) currentBucket = line;
    const isPersonnel = hasAny(line, RULES.personnelExpenseKeywords) || /기사급여|영양사|조리.*급여/.test(line);
    const allowed = hasAny(currentBucket, RULES.allowedPersonnelBuckets);
    if (isPersonnel && currentBucket && !allowed && !/세출합계|발행일/.test(line)) {
      const location = currentBucket.replace(/\s+/g, " ");
      const item = line.replace(/\s+/g, " ");
      suspicious.push({ item, location });
    }
  }
  const unique = Array.from(new Map(suspicious.map(x => [`${x.item}-${x.location}`, x])).values()).slice(0, 20);
  unique.forEach(x => results.push({ category: "오편성", item: x.item, status: "bad", message: `${x.item}를 ${x.location}에 편성`, excelAmount: null, pdfAmount: null }));
}

function hasAny(text, keywords) {
  const n = norm(text);
  return keywords.some(k => n.includes(norm(k)));
}

function buildDebug(groups, pdfItems) {
  return [
    "\n[엑셀 자동 집계]",
    `교원급여: ${fmt(groups.teacherPay.amount)}원 / ${groups.teacherPay.people}명`,
    `방과후교원급여: ${fmt(groups.afterSchoolTeacherPay.amount)}원 / ${groups.afterSchoolTeacherPay.people}명`,
    `직원급여: ${fmt(groups.staffPay.amount)}원 / ${groups.staffPay.people}명`,
    `교원수당 후보: ${groups.teacherAllowances.map(a => `${a.name}=${fmt(a.amount)}`).join(", ") || "없음"}`,
    `직원수당 후보: ${groups.staffAllowances.map(a => `${a.name}=${fmt(a.amount)}`).join(", ") || "없음"}`,
    "\n[PDF 산출기초 후보 일부]",
    pdfItems.slice(0, 40).join("\n")
  ];
}

function renderReport(report) {
  const counts = { ok: 0, info: 0, warn: 0, bad: 0 };
  report.results.forEach(r => counts[r.status]++);
  const summary = $("summary");
  summary.hidden = false;
  summary.innerHTML = `
    <div class="metric">정상<strong>${counts.ok}</strong></div>
    <div class="metric">통합편성<strong>${counts.info}</strong></div>
    <div class="metric">추가확인<strong>${counts.warn}</strong></div>
    <div class="metric">지적<strong>${counts.bad}</strong></div>
  `;
  const badgeText = { ok: "정상", info: "통합편성", warn: "추가확인", bad: "지적" };
  $("results").innerHTML = `
    <div class="table-wrap"><table>
      <thead><tr><th>구분</th><th>항목</th><th>엑셀 금액</th><th>PDF 금액</th><th>엑셀 인원</th><th>PDF 인원</th><th>판정</th><th>내용</th></tr></thead>
      <tbody>${report.results.map(r => `
        <tr>
          <td>${escapeHtml(r.category)}</td>
          <td>${escapeHtml(r.item)}</td>
          <td>${r.excelAmount == null ? "-" : fmt(r.excelAmount) + "원"}</td>
          <td>${r.pdfAmount == null ? "-" : fmt(r.pdfAmount) + "원"}</td>
          <td>${r.excelPeople ?? "-"}</td>
          <td>${r.pdfPeople ?? "-"}</td>
          <td><span class="badge ${r.status}">${badgeText[r.status]}</span></td>
          <td>${escapeHtml(r.message)}</td>
        </tr>`).join("")}</tbody>
    </table></div>`;
  $("debugLog").textContent = report.debug.join("\n");
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"]/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[m]));
}
