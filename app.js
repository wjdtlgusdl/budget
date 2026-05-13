/* global XLSX, pdfjsLib */

const RULES = {
  retirementSheetKeywords: ["퇴직", "적립", "충당"],
  personnelExpenseKeywords: ["인건비", "급여", "수당", "퇴직", "적립금", "상여금"],
  allowedPersonnelBuckets: ["교원인건비", "직원인건비", "그밖의인건비"],
  paymentStopHeaders: ["지급액계", "소득세", "주민세", "본인부담금", "공제액계", "실수령액"],
  basePayHeaders: ["본봉", "기본급"],
  allowanceHeaders: ["시간외", "직급보조", "관리업무", "기타수당", "자가운전", "연구활동", "식대", "정액급식", "명절", "스승", "방학", "성과", "상여"],
  excludedLegalBurdenKeywords: ["법정부담금", "4대보험", "사학연금", "국민연금", "건강보험", "고용보험", "산재보험", "장기요양", "보험료"],
  directAllowanceAliases: {
    "식대": ["교원정액급식비", "직원정액급식비", "정액급식비"],
    "명절휴가비": ["교원명절휴가비", "직원명절휴가비", "명절휴가비"],
    "스승의날상여금": ["교원스승의날상여금", "직원스승의날상여금", "스승의날상여금"],
    "방학휴가비": ["교원방학휴가비", "직원방학휴가비", "방학휴가비"],
    "성과상여금": ["교원성과상여금", "직원성과상여금", "성과상여금"]
  }
};

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const $ = (id) => document.getElementById(id);
const fmt = (n) => Number.isFinite(n) ? Math.round(n).toLocaleString("ko-KR") : "-";
const norm = (s) => String(s ?? "").replace(/\s+/g, "").trim();
const toNumber = (v) => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const cleaned = String(v ?? "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
};
const approxEqual = (a, b, tolerance = 1000) => Math.abs((a || 0) - (b || 0)) <= tolerance;

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
  const wb = XLSX.read(data, { type: "array", cellDates: false, cellFormula: true, raw: true });
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    return { name, matrix, text: matrix.flat().map(String).join(" ") };
  });
  const salarySheet = pickSalarySheet(sheets);
  if (!salarySheet) throw new Error("교직원보수일람표 시트를 찾지 못했습니다.");
  const retirementSheets = sheets.filter(s => hasAny(s.name + " " + s.text, RULES.retirementSheetKeywords));
  const salaryRows = inferSalaryRows(salarySheet.matrix);
  return { sheets, salarySheet, salaryRows, retirementSheets };
}

function pickSalarySheet(sheets) {
  return sheets.find(s => norm(s.name).includes("교직원보수일람표"))
    || sheets.find(s => norm(s.text).includes("교직원보수일람표"))
    || sheets.find(s => hasAny(s.name, ["교직원", "보수", "일람"]));
}

function inferSalaryRows(matrix) {
  const headerRowIndex = matrix.findIndex(row => row.some(c => norm(c).includes("본봉")) && row.some(c => norm(c).includes("성명")));
  if (headerRowIndex < 0) return [];

  const header1 = matrix[headerRowIndex] || [];
  const header2 = matrix[headerRowIndex + 1] || [];
  const headers = header1.map((h, i) => cleanHeader([h, header2[i]].filter(Boolean).join(" ")) || `열${i + 1}`);

  const roleCol = headers.findIndex(h => norm(h).includes("직명"));
  const nameCol = headers.findIndex(h => norm(h).includes("성명"));
  const baseCol = headers.findIndex(h => hasAny(h, RULES.basePayHeaders));
  const stopCol = headers.findIndex(h => hasAny(h, RULES.paymentStopHeaders));
  const lastPayCol = stopCol > 0 ? stopCol - 1 : headers.length - 1;

  const rows = [];
  for (let r = headerRowIndex + 2; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const marker = norm(row[0]);
    if (!marker || marker.includes("작성요령")) continue;
    if (marker.includes("소계") || marker.includes("합계")) continue;
    const role = String(row[roleCol] ?? "").trim();
    const name = String(row[nameCol] ?? "").trim();
    if (!role || !name) continue;
    const basePay = toNumber(row[baseCol]);
    const amounts = {};
    for (let c = 0; c <= lastPayCol; c++) {
      const value = toNumber(row[c]);
      if (value > 0) amounts[headers[c]] = (amounts[headers[c]] || 0) + value;
    }
    if (basePay <= 0 && Object.keys(amounts).length === 0) continue;
    rows.push({ rowNumber: r + 1, role, name, rowText: `${role} ${name}`, basePay, amounts, headers });
  }
  return rows;
}

function cleanHeader(s) {
  return String(s ?? "")
    .replace(/\([^)]*\)/g, "")
    .replace(/[A-Z]=[^\s]+/g, "")
    .replace(/\s+/g, "")
    .trim();
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
  const lines = pages.flatMap(p => p.lines.map(line => ({ page: p.page, line })));
  const fullText = pages.map(p => p.text).join("\n");
  return { pages, lines, fullText, compactText: norm(fullText), budgetRows: extractBudgetRows(lines), formulaRows: extractFormulaRows(lines) };
}

function extractBudgetRows(lines) {
  const rows = [];
  for (const { page, line } of lines) {
    const cleaned = line.replace(/\s+/g, " ").trim();
    const m = cleaned.match(/^(.+?)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)$/);
    if (!m) continue;
    const label = m[1].trim();
    if (!/[가-힣]/.test(label) || /예산구분|발행일|보조금|수익자/.test(label)) continue;
    rows.push({ page, label, compactLabel: norm(label), grant: toNumber(m[2]) * 1000, parent: toNumber(m[3]) * 1000, other: toNumber(m[4]) * 1000, total: toNumber(m[5]) * 1000, raw: line });
  }
  return rows;
}

function extractFormulaRows(lines) {
  const rows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].line;
    const m = line.match(/([0-9,]+)\s*원\s*\*\s*([^=]+?)\s*=\s*([0-9,]+)/);
    if (!m) continue;
    const near = [];
    for (let j = Math.max(0, i - 2); j <= Math.min(lines.length - 1, i + 3); j++) near.push(lines[j].line);
    const peopleMatches = [...line.matchAll(/\*\s*([0-9]+)\s*명/g)].map(x => Number(x[1]));
    rows.push({ page: lines[i].page, line, windowText: near.join(" "), compactWindow: norm(near.join(" ")), unit: toNumber(m[1]), total: toNumber(m[3]), people: peopleMatches.length ? Math.max(...peopleMatches) : 0 });
  }
  return rows;
}

function analyze(excel, pdf) {
  const results = [];
  const debug = [];
  const groups = buildExcelGroups(excel.salaryRows);

  debug.push(`선택된 보수 시트: ${excel.salarySheet?.name || "확인 불가"}`);
  debug.push(`보수 대상 행: ${excel.salaryRows.length}건`);
  debug.push(`퇴직 관련 시트 후보: ${excel.retirementSheets.map(s => s.name).join(", ") || "없음"}`);
  debug.push(`PDF 예산 과목 행: ${pdf.budgetRows.length}건 / 산출기초 행: ${pdf.formulaRows.length}건`);

  comparePay(results, "교원급여", groups.teacherPay, getPdfItemAmount(pdf, ["교원급여"]), getPdfPeople(pdf, ["교원급여"], ["방과후"]), true);
  comparePay(results, "방과후교원급여", groups.afterSchoolTeacherPay, getPdfFormulaAmount(pdf, ["방과후교원급여", "방과후급여"]), getPdfPeople(pdf, ["방과후교원급여", "방과후급여"]), true);
  comparePay(results, "직원급여", groups.staffPay, getPdfItemAmount(pdf, ["직원급여"]), getPdfPeople(pdf, ["직원급여"], []), false);

  checkAllowances(results, groups.teacherAllowances, pdf, "교원수당", "교원");
  checkAllowances(results, groups.staffAllowances, pdf, "직원수당", "직원");
  checkRetirement(results, excel, pdf);
  checkMisclassifiedPersonnel(results, pdf);

  return { results, debug: debug.concat(buildDebug(groups, pdf)) };
}

function buildExcelGroups(rows) {
  const regularTeachers = rows.filter(r => isTeacher(r.role) && !norm(r.role).includes("방과후"));
  const afterSchoolTeachers = rows.filter(r => norm(r.role).includes("방과후"));
  const staff = rows.filter(r => !regularTeachers.includes(r) && !afterSchoolTeachers.includes(r));

  return {
    teacherPay: sumBase(regularTeachers),
    afterSchoolTeacherPay: sumBase(afterSchoolTeachers),
    staffPay: sumBase(staff),
    teacherAllowances: collectAllowances(regularTeachers),
    staffAllowances: collectAllowances(staff),
    staffRoles: staff.map(r => r.role)
  };
}

function isTeacher(role) {
  const r = norm(role);
  return r.includes("원장") || r.includes("원감") || r.includes("교사") || r.includes("교원") || r.includes("담임");
}

function sumBase(rows) {
  return { amount: rows.reduce((a, r) => a + (r.basePay || 0), 0), people: rows.filter(r => (r.basePay || 0) > 0).length, rows };
}

function collectAllowances(rows) {
  const map = new Map();
  rows.forEach(row => {
    Object.entries(row.amounts).forEach(([header, value]) => {
      if (value <= 0) return;
      if (isLegalBurden(header)) return;
      if (hasAny(header, RULES.basePayHeaders)) return;
      if (!hasAny(header, RULES.allowanceHeaders)) return;
      const label = canonicalAllowanceName(header);
      map.set(label, (map.get(label) || 0) + value);
    });
  });
  return [...map.entries()].map(([name, amount]) => ({ name, amount })).filter(x => x.amount > 0);
}

function canonicalAllowanceName(header) {
  const h = norm(header);
  if (h.includes("시간외")) return "시간외수당";
  if (h.includes("직급보조")) return "직급보조비";
  if (h.includes("관리업무")) return "관리업무수당";
  if (h.includes("기타수당")) return "기타수당";
  if (h.includes("자가운전")) return "자가운전보조금";
  if (h.includes("연구활동")) return "연구활동비";
  if (h.includes("식대") || h.includes("정액급식")) return "식대";
  if (h.includes("명절")) return "명절휴가비";
  if (h.includes("스승")) return "스승의날상여금";
  if (h.includes("방학")) return "방학휴가비";
  if (h.includes("성과")) return "성과상여금";
  return header;
}

function getPdfItemAmount(pdf, aliases) {
  const aliasNorms = aliases.map(norm);
  const row = pdf.budgetRows.find(r => aliasNorms.includes(r.compactLabel));
  return row ? row.total : 0;
}

function getPdfFormulaAmount(pdf, aliases, excludeAliases = []) {
  const aliasNorms = aliases.map(norm);
  const excludeNorms = excludeAliases.map(norm);
  const rows = pdf.formulaRows.filter(r => aliasNorms.some(a => r.compactWindow.includes(a)) && !excludeNorms.some(e => r.compactWindow.includes(e)));
  return rows.reduce((a, r) => a + r.total, 0);
}

function getPdfPeople(pdf, aliases, excludeAliases = []) {
  const aliasNorms = aliases.map(norm);
  const excludeNorms = excludeAliases.map(norm);
  const rows = pdf.formulaRows.filter(r => aliasNorms.some(a => r.compactWindow.includes(a)) && !excludeNorms.some(e => r.compactWindow.includes(e)) && r.people > 0);
  if (!rows.length) return 0;
  return Math.max(...rows.map(r => r.people));
}

function comparePay(results, label, excelGroup, pdfAmount, pdfPeople, checkPeople) {
  const amountDiff = excelGroup.amount - pdfAmount;
  const peopleDiff = excelGroup.people - pdfPeople;
  let status = "ok", message = "금액 일치";
  if (excelGroup.amount === 0 && pdfAmount === 0) {
    status = "warn"; message = "자동 추출 실패 또는 항목 없음";
  } else if (!approxEqual(excelGroup.amount, pdfAmount)) {
    status = "bad"; message = `금액 차이 ${fmt(amountDiff)}원`;
  }
  if (checkPeople && pdfPeople > 0 && peopleDiff !== 0) {
    status = "bad";
    message += `${message ? " / " : ""}인원수 불일치: PDF ${pdfPeople}명, 엑셀 ${excelGroup.people}명`;
  }
  results.push({ category: "보수", item: label, excelAmount: excelGroup.amount, pdfAmount, excelPeople: excelGroup.people, pdfPeople, status, message });
}

function checkAllowances(results, allowances, pdf, bucketName, prefix) {
  const missing = [];
  const direct = [];
  allowances.forEach(a => {
    const aliases = allowanceAliasesFor(a.name, prefix);
    const amount = getPdfFormulaAmount(pdf, aliases);
    if (amount > 0 && approxEqual(amount, a.amount)) direct.push({ ...a, pdfAmount: amount });
    else missing.push(a);
  });

  direct.forEach(a => results.push({ category: "수당", item: `${prefix} ${a.name}`, excelAmount: a.amount, pdfAmount: a.pdfAmount, status: "ok", message: "개별 편성 일치" }));

  const missingSum = missing.reduce((a, b) => a + b.amount, 0);
  if (missing.length === 0) return;
  const bucketAmount = getPdfFormulaAmount(pdf, [bucketName]);
  if (bucketAmount && approxEqual(bucketAmount, missingSum)) {
    results.push({ category: "수당", item: bucketName, excelAmount: missingSum, pdfAmount: bucketAmount, status: "info", message: `통합편성(${missing.map(x => `${x.name} ${fmt(x.amount)}원`).join(" + ")} = ${fmt(missingSum)}원)` });
  } else {
    results.push({ category: "수당", item: bucketName, excelAmount: missingSum, pdfAmount: bucketAmount, status: "warn", message: `추가확인필요: PDF에서 개별 확인되지 않은 수당 ${missing.length}건의 합계와 ${bucketName} 산출기초 금액이 일치하지 않습니다.` });
  }
}

function allowanceAliasesFor(name, prefix) {
  const base = RULES.directAllowanceAliases[name] || [name];
  return Array.from(new Set(base.flatMap(x => [x, `${prefix}${x}`, `${prefix}${name}`])));
}

function checkRetirement(results, excel, pdf) {
  const hasRetirementAmount = excel.retirementSheets.some(s => s.matrix.flat().some(c => toNumber(c) > 0));
  if (!hasRetirementAmount) return;
  const found = hasAny(pdf.fullText, ["퇴직", "퇴직금", "퇴직적립", "퇴직급여충당"]);
  results.push({ category: "퇴직금", item: "퇴직적립금 편성 여부", status: found ? "ok" : "bad", message: found ? "퇴직 관련 편성 항목 확인" : "퇴직금 적립금 미편성 의심", excelAmount: null, pdfAmount: null });
}

function checkMisclassifiedPersonnel(results, pdf) {
  const suspicious = [];
  const importantNonPersonnelBuckets = ["일반급식비간식비", "통학차량이용비", "특별급식비간식비", "방과후교육활동비", "시설비", "수용비"];
  let currentBudgetRow = "";

  for (let i = 0; i < pdf.lines.length; i++) {
    const line = pdf.lines[i].line;
    const budget = pdf.budgetRows.find(r => r.raw === line);
    if (budget) currentBudgetRow = budget.label;

    const window = pdf.lines.slice(Math.max(0, i - 2), Math.min(pdf.lines.length, i + 4)).map(x => x.line).join(" ");
    const nWindow = norm(window);
    if (isLegalBurden(nWindow)) continue;
    const isPersonnel = hasAny(nWindow, RULES.personnelExpenseKeywords) || /기사급여|영양사|조리.*급여/.test(nWindow);
    const allowed = hasAny(currentBudgetRow, RULES.allowedPersonnelBuckets);
    const inWrongBucket = hasAny(currentBudgetRow, importantNonPersonnelBuckets) || (!allowed && currentBudgetRow && !/인건비/.test(currentBudgetRow));
    if (isPersonnel && inWrongBucket && !/세출합계|발행일|예산구분/.test(line)) {
      const item = extractPersonnelLabel(window);
      if (item) suspicious.push({ item, location: currentBudgetRow });
    }
  }

  const unique = Array.from(new Map(suspicious.map(x => [`${x.item}-${x.location}`, x])).values()).slice(0, 20);
  unique.forEach(x => results.push({ category: "오편성", item: x.item, status: "bad", message: `${x.item}를 ${x.location}에 편성`, excelAmount: null, pdfAmount: null }));
}

function extractPersonnelLabel(text) {
  const compact = norm(text);
  if (isLegalBurden(text)) return "";
  const candidates = ["영양사인건비", "영양사급여", "차량기사급여", "조리직원급여", "조리원급여", "교원급여", "직원급여", "인건비", "수당"];
  return candidates.find(c => compact.includes(norm(c))) || "인건비성 항목";
}

function isLegalBurden(text) {
  return hasAny(text, RULES.excludedLegalBurdenKeywords);
}

function hasAny(text, keywords) {
  const n = norm(text);
  return keywords.some(k => n.includes(norm(k)));
}

function buildDebug(groups, pdf) {
  return [
    "\n[엑셀 자동 집계]",
    `교원급여: ${fmt(groups.teacherPay.amount)}원 / ${groups.teacherPay.people}명`,
    `방과후교원급여: ${fmt(groups.afterSchoolTeacherPay.amount)}원 / ${groups.afterSchoolTeacherPay.people}명`,
    `직원급여: ${fmt(groups.staffPay.amount)}원 / ${groups.staffPay.people}명`,
    `교원수당 후보: ${groups.teacherAllowances.map(a => `${a.name}=${fmt(a.amount)}`).join(", ") || "없음"}`,
    `직원수당 후보: ${groups.staffAllowances.map(a => `${a.name}=${fmt(a.amount)}`).join(", ") || "없음"}`,
    "\n[PDF 예산 과목 일부]",
    pdf.budgetRows.slice(0, 35).map(r => `${r.label}: ${fmt(r.total)}원`).join("\n"),
    "\n[PDF 산출기초 일부]",
    pdf.formulaRows.slice(0, 45).map(r => `${r.line} / 근처: ${r.windowText}`).join("\n")
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
