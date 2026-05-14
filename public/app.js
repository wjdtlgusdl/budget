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

  debug.push(`선택된 보수 시트: ${excel.salarySheet?.name || "확인 불가"}`);
  debug.push(`퇴직 관련 시트 후보: ${excel.retirementSheets.map(s => s.name).join(", ") || "없음"}`);

  const salaryModel = buildSalaryModel(excel.salarySheet?.matrix || []);
  const pdfModel = buildPdfModel(pdf);

  debug.push(`교원 구간: ${salaryModel.teacherRows.length}행`);
  debug.push(`직원 구간: ${salaryModel.staffRows.length}행`);
  debug.push(`수당 열: ${salaryModel.allowanceHeaders.map(h => h.name).join(", ") || "없음"}`);
  debug.push(`자가운전보조금 합계: ${fmt(salaryModel.selfDrivingAllowance)}원`);

  checkSelfDrivingAllowance(results, salaryModel, pdfModel);
  checkVehicleDriverMisclassification(results, pdfModel);
  checkRetirementFocused(results, excel, pdfModel);

  return { results, debug };
}

function buildSalaryModel(matrix) {
  const headerRowIndex = matrix.findIndex(row => row.some(c => norm(c).includes("직명")) && row.some(c => hasAny(c, ["본봉", "기본급"])));
  const headerRow1 = headerRowIndex >= 0 ? matrix[headerRowIndex] : [];
  const headerRow2 = headerRowIndex >= 0 ? matrix[headerRowIndex + 1] || [] : [];
  const maxCols = Math.max(headerRow1.length, headerRow2.length, ...matrix.map(r => r.length));
  const headers = Array.from({ length: maxCols }, (_, i) => ({
    index: i,
    name: String([headerRow1[i], headerRow2[i]].filter(Boolean).join(" ")).replace(/\s+/g, " ").trim()
  }));

  const baseCol = headers.find(h => hasAny(h.name, ["본봉", "기본급"]))?.index ?? -1;
  const payTotalCol = headers.find(h => norm(h.name).includes("지급액계"))?.index ?? -1;
  const allowanceHeaders = headers.filter(h => baseCol >= 0 && payTotalCol >= 0 && h.index > baseCol && h.index < payTotalCol && h.name);

  const teacherSubtotal = matrix.findIndex(row => row.some(c => /소계\s*\(\s*교원\s*\)/.test(String(c))));
  const staffSubtotal = matrix.findIndex(row => row.some(c => /소계\s*\(\s*(직원|일반직)\s*\)/.test(String(c))));
  const dataStart = headerRowIndex >= 0 ? headerRowIndex + 2 : 0;

  const teacherRows = teacherSubtotal >= 0 ? matrix.slice(dataStart, teacherSubtotal).filter(isRealSalaryRow) : [];
  const staffRows = teacherSubtotal >= 0 && staffSubtotal >= 0 ? matrix.slice(teacherSubtotal + 1, staffSubtotal).filter(isRealSalaryRow) : [];
  const allRows = teacherRows.concat(staffRows);

  const selfDrivingCols = allowanceHeaders.filter(h => hasAny(h.name, ["자가운전", "자가 운전"]));
  const selfDrivingAllowance = sumRowsByColumns(allRows, selfDrivingCols.map(h => h.index));
  const teacherSelfDrivingAllowance = sumRowsByColumns(teacherRows, selfDrivingCols.map(h => h.index));
  const staffSelfDrivingAllowance = sumRowsByColumns(staffRows, selfDrivingCols.map(h => h.index));

  return { headerRowIndex, headers, baseCol, payTotalCol, allowanceHeaders, teacherRows, staffRows, selfDrivingAllowance, teacherSelfDrivingAllowance, staffSelfDrivingAllowance };
}

function isRealSalaryRow(row) {
  const first = String(row[0] ?? "").trim();
  const job = String(row[1] ?? "").trim();
  return !!job && !/소계|합계|작성요령/.test(first + job);
}

function sumRowsByColumns(rows, cols) {
  return rows.reduce((sum, row) => sum + cols.reduce((a, c) => a + toNumber(row[c]), 0), 0);
}

function buildPdfModel(pdf) {
  const fullText = pdf.fullText;
  const compactText = norm(fullText);
  return { fullText, compactText, lines: fullText.split(/\n+/).map(s => s.trim()).filter(Boolean) };
}

function checkSelfDrivingAllowance(results, salaryModel, pdfModel) {
  if (salaryModel.selfDrivingAllowance <= 0) return;
  const hasPdfSelfDriving = hasAny(pdfModel.fullText, ["자가운전보조금", "자가 운전 보조금", "자가운전"]);
  if (!hasPdfSelfDriving) {
    const detail = [];
    if (salaryModel.teacherSelfDrivingAllowance > 0) detail.push(`교원 ${fmt(salaryModel.teacherSelfDrivingAllowance)}원`);
    if (salaryModel.staffSelfDrivingAllowance > 0) detail.push(`직원 ${fmt(salaryModel.staffSelfDrivingAllowance)}원`);
    results.push({
      category: "미편성",
      item: "자가운전보조금",
      excelAmount: salaryModel.selfDrivingAllowance,
      pdfAmount: 0,
      status: "bad",
      message: `자가운전보조금 예산 미편성${detail.length ? ` (${detail.join(" + ")} = ${fmt(salaryModel.selfDrivingAllowance)}원)` : ""}`
    });
  }
}

function checkVehicleDriverMisclassification(results, pdfModel) {
  const compact = pdfModel.compactText;
  const start = compact.indexOf(norm("통학차량이용비"));
  if (start < 0) return;
  const endCandidates = ["특별급식비간식비", "적립금", "시설설비비품비"].map(k => compact.indexOf(norm(k), start + 1)).filter(i => i > start);
  const end = endCandidates.length ? Math.min(...endCandidates) : start + 1200;
  const section = compact.slice(start, end);
  const m = section.match(/차량기사급여.*?=([0-9,]+)원?/);
  if (m) {
    const amount = toNumber(m[1]);
    results.push({
      category: "오편성",
      item: "차량기사급여",
      excelAmount: null,
      pdfAmount: amount,
      status: "bad",
      message: `차량기사급여 ${fmt(amount)}원을 통학차량이용비의 산출내역에 편성`
    });
  }
}

function checkRetirementFocused(results, excel, pdfModel) {
  const hasRetirementSheetAmount = excel.retirementSheets.some(s => {
    const text = norm(s.name + " " + s.text);
    if (!text.includes("퇴직")) return false;
    return s.matrix.flat().some(c => toNumber(c) > 0);
  });
  if (!hasRetirementSheetAmount) {
    results.push({ category: "퇴직금", item: "퇴직적립금", status: "warn", message: "퇴직 관련 시트의 적립금액을 자동 확인하지 못했습니다.", excelAmount: null, pdfAmount: null });
    return;
  }
  const hasPdfRetirement = hasAny(pdfModel.fullText, ["퇴직금", "퇴직적립", "퇴직 급여", "퇴직급여", "퇴직충당", "퇴직 충당"]);
  if (!hasPdfRetirement) {
    results.push({ category: "미편성", item: "퇴직적립금", status: "bad", message: "퇴직적립금 미편성", excelAmount: null, pdfAmount: 0 });
  } else {
    results.push({ category: "퇴직금", item: "퇴직적립금", status: "ok", message: "퇴직 관련 편성 항목 확인", excelAmount: null, pdfAmount: null });
  }
}

function hasAny(text, keywords) {
  const n = norm(text);
  return keywords.some(k => n.includes(norm(k)));
}

function buildDebug(groups, pdfItems) {
  return [];
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
