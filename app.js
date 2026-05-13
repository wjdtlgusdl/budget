/* global XLSX, pdfjsLib */

const RULES = {
  retirementSheetKeywords: ["퇴직", "적립", "충당"],
  allowedPersonnelBuckets: ["교원인건비", "직원인건비", "그밖의인건비"],
  basePayHeaders: ["본봉", "기본급"],
  allowanceHeaders: ["시간외", "직급보조", "관리업무", "기타수당", "자가운전", "연구활동", "식대", "정액급식", "명절", "스승", "방학", "성과", "상여"],
  excludedLegalBurdenKeywords: ["법정부담금", "4대보험", "사학연금", "국민연금", "건강보험", "고용보험", "산재보험", "장기요양", "사회보험", "보험료"],
  directAllowanceAliases: {
    "식대": ["교원정액급식비", "직원정액급식비", "정액급식비"],
    "명절휴가비": ["교원명절휴가비", "직원명절휴가비", "명절휴가비"],
    "스승의날상여금": ["교원스승의날상여금", "직원스승의날상여금", "스승의날상여금"],
    "방학휴가비": ["교원방학휴가비", "직원방학휴가비", "방학휴가비"],
    "성과상여금": ["교원성과상여금", "직원성과상여금", "성과상여금"]
  },
  specificPersonnelItems: [
    "영양사인건비", "영양사급여", "영양사", "조리직원급여", "조리원급여", "조리사급여",
    "차량기사급여", "차량기사", "보조교사급여", "환경미화원급여"
  ]
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
  const wb = XLSX.read(data, { type: "array", cellDates: false, cellFormula: false, raw: true, blankrows: false });
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const matrix = worksheetToMatrix(ws);
    return { name, matrix, text: matrix.flat().map(String).join(" ") };
  });
  const salarySheet = pickSalarySheet(sheets);
  if (!salarySheet) throw new Error("교직원보수일람표 시트를 찾지 못했습니다.");
  const retirementSheets = sheets.filter(s => norm(s.name + " " + s.text).includes("퇴직"));
  const salaryData = inferSalaryData(salarySheet.matrix);
  // v7: 사용자 기준에 따라 엑셀은 보수기준 시트와 퇴직 관련 시트만 활용한다.
  // 월급여/비월정수당 등 다른 시트는 검토에서 완전히 제외한다.
  return { sheets, salarySheet, salaryRows: salaryData.rows, salaryData, retirementSheets };
}


function worksheetToMatrix(ws) {
  const ref = ws["!ref"];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  const matrix = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell) { row.push(""); continue; }
      // SheetJS가 수식 셀을 읽을 때 표시문자(w) 또는 캐시값(v) 중 숫자를 우선한다.
      if (typeof cell.v === "number") row.push(cell.v);
      else if (cell.v != null && String(cell.v).trim() !== "") row.push(cell.v);
      else if (cell.w != null) row.push(cell.w);
      else row.push("");
    }
    matrix.push(row);
  }
  return matrix;
}

function pickSalarySheet(sheets) {
  return sheets.find(s => norm(s.name).includes("교직원보수일람표"))
    || sheets.find(s => norm(s.text).includes("교직원보수일람표"))
    || sheets.find(s => hasAny(s.name, ["교직원", "보수", "일람"]));
}

function inferSalaryData(matrix) {
  // 사용자가 알려준 구조를 우선 적용:
  // 직명 헤더 아래부터 소계(교원) 전까지 = 교원, 그 아래부터 소계(직원/일반직) 전까지 = 직원.
  // 본봉/기본급 열 바로 다음부터 지급액계 바로 전까지 = 수당.
  const headerRowIndex = matrix.findIndex(row => row.some(c => norm(c).includes("직명")) && row.some(c => hasAny(c, RULES.basePayHeaders)));
  if (headerRowIndex < 0) return emptySalaryData("보수표 헤더를 찾지 못했습니다.");

  const headerRows = [matrix[headerRowIndex] || [], matrix[headerRowIndex + 1] || []];
  const headers = buildMergedHeaders(headerRows);
  const roleCol = findHeaderIndex(headers, ["직명"], 1);
  const nameCol = findHeaderIndex(headers, ["성명", "성 명"], 2);
  const baseCol = findHeaderIndex(headers, RULES.basePayHeaders, 4);
  const payTotalCol = findHeaderIndex(headers, ["지급액계", "지급액", "급여계"], 16);
  const allowanceStart = baseCol + 1;
  const allowanceEnd = payTotalCol > allowanceStart ? payTotalCol - 1 : Math.max(allowanceStart, 15);

  const teacherSubtotalIndex = findSubtotalIndex(matrix, headerRowIndex + 1, ["교원"]);
  const staffSubtotalIndex = findSubtotalIndex(matrix, teacherSubtotalIndex >= 0 ? teacherSubtotalIndex + 1 : headerRowIndex + 1, ["직원", "일반직"]);

  const teacherStart = skipHeaderRows(matrix, headerRowIndex + 1);
  const teacherEnd = teacherSubtotalIndex >= 0 ? teacherSubtotalIndex - 1 : -1;
  const staffStart = teacherSubtotalIndex >= 0 ? teacherSubtotalIndex + 1 : -1;
  const staffEnd = staffSubtotalIndex >= 0 ? staffSubtotalIndex - 1 : -1;

  const rows = [];
  if (teacherEnd >= teacherStart) rows.push(...readPersonRows(matrix, teacherStart, teacherEnd, "교원", { roleCol, nameCol, baseCol, allowanceStart, allowanceEnd, headers }));
  if (staffEnd >= staffStart) rows.push(...readPersonRows(matrix, staffStart, staffEnd, "직원", { roleCol, nameCol, baseCol, allowanceStart, allowanceEnd, headers }));

  const teacherSubtotal = readSubtotal(matrix[teacherSubtotalIndex], { baseCol, allowanceStart, allowanceEnd, headers, label: "소계(교원)" });
  const staffSubtotal = readSubtotal(matrix[staffSubtotalIndex], { baseCol, allowanceStart, allowanceEnd, headers, label: "소계(직원/일반직)" });

  const teacherRows = rows.filter(r => r.group === "교원");
  const staffRows = rows.filter(r => r.group === "직원");
  return {
    rows,
    headers,
    meta: { headerRowIndex, roleCol, nameCol, baseCol, allowanceStart, allowanceEnd, payTotalCol },
    sections: {
      teacher: {
        label: "교원",
        startRow: teacherStart + 1,
        endRow: teacherEnd + 1,
        subtotalRow: teacherSubtotalIndex + 1,
        rows: teacherRows,
        subtotal: teacherSubtotal,
        people: countPeopleRows(teacherRows)
      },
      staff: {
        label: "직원",
        startRow: staffStart + 1,
        endRow: staffEnd + 1,
        subtotalRow: staffSubtotalIndex + 1,
        rows: staffRows,
        subtotal: staffSubtotal,
        people: countPeopleRows(staffRows)
      }
    },
    warnings: []
  };
}

function emptySalaryData(message) {
  return { rows: [], headers: [], meta: {}, sections: { teacher: blankSection("교원"), staff: blankSection("직원") }, warnings: [message] };
}
function blankSection(label) { return { label, startRow: 0, endRow: 0, subtotalRow: 0, rows: [], subtotal: { amount: 0, allowances: [] }, people: 0 }; }

function buildMergedHeaders(headerRows) {
  const maxCols = Math.max(...headerRows.map(r => r.length), 24);
  const headers = [];
  for (let c = 0; c < maxCols; c++) {
    const parts = headerRows.map(r => r[c]).filter(v => String(v ?? "").trim() !== "");
    headers[c] = cleanHeader(parts.join(" ")) || `열${c + 1}`;
  }
  return headers;
}

function findSubtotalIndex(matrix, fromRow, keywords) {
  for (let r = Math.max(0, fromRow); r < matrix.length; r++) {
    const rowText = norm((matrix[r] || []).slice(0, 5).join(" "));
    if (!rowText.includes("소계")) continue;
    if (keywords.some(k => rowText.includes(norm(k)))) return r;
  }
  return -1;
}

function skipHeaderRows(matrix, fromRow) {
  let r = fromRow;
  while (r < matrix.length) {
    const rowText = norm((matrix[r] || []).join(" "));
    if (!rowText || rowText.includes("직명") || rowText.includes("성명") || rowText.includes("사학연금") || rowText.includes("사회보험")) { r++; continue; }
    break;
  }
  return r;
}

function readPersonRows(matrix, start, end, group, cols) {
  const rows = [];
  for (let r = start; r <= end && r < matrix.length; r++) {
    const row = matrix[r] || [];
    const rowText = norm(row.join(" "));
    if (!rowText || rowText.includes("소계") || rowText.includes("합계") || rowText.includes("작성요령")) continue;
    const serial = String(row[0] ?? "").trim();
    // 직명/성명이 기관별 수식 때문에 깨져도 일련번호가 있으면 인원으로 인정한다.
    if (!/^\d+$/.test(serial)) continue;
    const role = String(row[cols.roleCol] ?? "").trim();
    const name = String(row[cols.nameCol] ?? "").trim();
    const basePay = toNumber(row[cols.baseCol]);
    const allowances = readAllowancesFromRow(row, cols);
    rows.push({ rowNumber: r + 1, group, serial: Number(serial), role, name, basePay, allowances, raw: row });
  }
  return rows;
}

function readSubtotal(row, cols) {
  if (!row) return { amount: 0, allowances: [], source: "없음" };
  const amount = toNumber(row[cols.baseCol]);
  const allowances = [];
  for (let c = cols.allowanceStart; c <= cols.allowanceEnd; c++) {
    const header = canonicalAllowanceName(cols.headers[c] || `열${c + 1}`);
    if (!header || isLegalBurden(header)) continue;
    const value = toNumber(row[c]);
    if (value > 0) allowances.push({ name: header, amount: value, col: c + 1 });
  }
  return { amount, allowances, source: cols.label };
}

function readAllowancesFromRow(row, cols) {
  const allowances = {};
  for (let c = cols.allowanceStart; c <= cols.allowanceEnd; c++) {
    const header = canonicalAllowanceName(cols.headers[c] || `열${c + 1}`);
    if (!header || isLegalBurden(header)) continue;
    const value = toNumber(row[c]);
    if (value > 0) allowances[header] = (allowances[header] || 0) + value;
  }
  return allowances;
}

function countPeopleRows(rows) { return rows.length; }

function findHeaderIndex(headers, keys, fallback) {
  const idx = headers.findIndex(h => hasAny(h, keys));
  return idx >= 0 ? idx : fallback;
}

function cleanHeader(s) {
  return String(s ?? "")
    .replace(/\([^)]*\)/g, "")
    .replace(/[A-Z]=[^\s]+/g, "")
    .replace(/\n/g, "")
    .replace(/\s+/g, "")
    .trim();
}


function pickMonthlySheet(sheets) {
  return sheets.find(s => norm(s.name).includes("월급여"))
    || sheets.find(s => norm(s.name).includes("급여대장"))
    || sheets.find(s => norm(s.text).includes("교직원월급여대장"));
}

function inferMonthlyData(matrix) {
  const headerRowIndex = matrix.findIndex(row => row.some(c => norm(c).includes("직명")) && row.some(c => hasAny(c, RULES.basePayHeaders)));
  if (headerRowIndex < 0) return emptySalaryData("월급여 헤더를 찾지 못했습니다.");
  const headers = buildMergedHeaders([matrix[headerRowIndex] || []]);
  const roleCol = findHeaderIndex(headers, ["직명"], 1);
  const nameCol = findHeaderIndex(headers, ["성명", "성 명"], 2);
  const monthCol = findHeaderIndex(headers, ["근무월수", "월수"], 4);
  const baseCol = findHeaderIndex(headers, RULES.basePayHeaders, 5);
  const payTotalCol = findHeaderIndex(headers, ["지급액계", "지급액", "급여계"], 15);
  const allowanceStart = baseCol + 1;
  const allowanceEnd = payTotalCol > allowanceStart ? payTotalCol - 1 : Math.max(allowanceStart, 14);

  const teacherSubtotalIndex = findSubtotalIndex(matrix, headerRowIndex + 1, ["교원"]);
  const staffSubtotalIndex = findSubtotalIndex(matrix, teacherSubtotalIndex >= 0 ? teacherSubtotalIndex + 1 : headerRowIndex + 1, ["직원", "일반직"]);
  const teacherStart = skipHeaderRows(matrix, headerRowIndex + 1);
  const teacherEnd = teacherSubtotalIndex >= 0 ? teacherSubtotalIndex - 1 : -1;
  const staffStart = teacherSubtotalIndex >= 0 ? teacherSubtotalIndex + 1 : -1;
  const staffEnd = staffSubtotalIndex >= 0 ? staffSubtotalIndex - 1 : -1;
  const cols = { roleCol, nameCol, monthCol, baseCol, allowanceStart, allowanceEnd, headers };
  const rows = [];
  if (teacherEnd >= teacherStart) rows.push(...readMonthlyRows(matrix, teacherStart, teacherEnd, "교원", cols));
  if (staffEnd >= staffStart) rows.push(...readMonthlyRows(matrix, staffStart, staffEnd, "직원", cols));
  return {
    rows,
    headers,
    meta: { headerRowIndex, roleCol, nameCol, monthCol, baseCol, allowanceStart, allowanceEnd, payTotalCol },
    sections: {
      teacher: { label: "교원", startRow: teacherStart + 1, endRow: teacherEnd + 1, subtotalRow: teacherSubtotalIndex + 1, rows: rows.filter(r => r.group === "교원"), subtotal: readMonthlySubtotal(matrix[teacherSubtotalIndex], cols), people: rows.filter(r => r.group === "교원").length },
      staff: { label: "직원", startRow: staffStart + 1, endRow: staffEnd + 1, subtotalRow: staffSubtotalIndex + 1, rows: rows.filter(r => r.group === "직원"), subtotal: readMonthlySubtotal(matrix[staffSubtotalIndex], cols), people: rows.filter(r => r.group === "직원").length }
    },
    warnings: []
  };
}

function readMonthlyRows(matrix, start, end, group, cols) {
  const rows = [];
  for (let r = start; r <= end && r < matrix.length; r++) {
    const row = matrix[r] || [];
    const rowText = norm(row.join(" "));
    if (!rowText || rowText.includes("소계") || rowText.includes("합계") || rowText.includes("작성요령")) continue;
    const serial = String(row[0] ?? "").trim();
    if (!/^\d+$/.test(serial)) continue;
    const role = String(row[cols.roleCol] ?? "").trim();
    const name = String(row[cols.nameCol] ?? "").trim();
    const months = toNumber(row[cols.monthCol]) || 12;
    const baseMonthly = toNumber(row[cols.baseCol]);
    const allowances = {};
    for (let c = cols.allowanceStart; c <= cols.allowanceEnd; c++) {
      const header = canonicalAllowanceName(cols.headers[c] || `열${c + 1}`);
      if (!header || isLegalBurden(header)) continue;
      const value = toNumber(row[c]);
      if (value > 0) allowances[header] = (allowances[header] || 0) + (value * months);
    }
    rows.push({ rowNumber: r + 1, group, serial: Number(serial), role, name, months, basePay: baseMonthly * months, baseMonthly, allowances, raw: row });
  }
  return rows;
}

function readMonthlySubtotal(row, cols) {
  if (!row) return { amount: 0, allowances: [], source: "없음" };
  // 월급여 소계는 월액 기준이므로 총액 검토에는 보수일람표 소계를 우선 사용한다.
  const allowances = [];
  for (let c = cols.allowanceStart; c <= cols.allowanceEnd; c++) {
    const header = canonicalAllowanceName(cols.headers[c] || `열${c + 1}`);
    if (!header || isLegalBurden(header)) continue;
    const value = toNumber(row[c]);
    if (value > 0) allowances.push({ name: header, amount: value * 12, col: c + 1 });
  }
  return { amount: toNumber(row[cols.baseCol]) * 12, allowances, source: "월급여 소계" };
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
      if (!linesMap.has(item.y)) linesMap.set(item.y, []);
      linesMap.get(item.y).push(item);
    });
    const lines = [...linesMap.entries()]
      .sort((a, b) => b[0] - a[0])
      .map(([, arr]) => arr.sort((a, b) => a.x - b.x).map(x => x.str).join(" ").trim())
      .filter(Boolean);
    pages.push({ page: i, lines, text: lines.join("\n") });
  }
  const lines = pages.flatMap(p => p.lines.map(line => ({ page: p.page, line })));
  const fullText = pages.map(p => p.text).join("\n");
  const budgetRows = extractBudgetRows(lines);
  const formulaRows = extractFormulaRows(lines, budgetRows);
  return { pages, lines, fullText, compactText: norm(fullText), budgetRows, formulaRows };
}

function extractBudgetRows(lines) {
  const rows = [];
  for (let i = 0; i < lines.length; i++) {
    const { page, line } = lines[i];
    const cleaned = line.replace(/\s+/g, " ").trim();
    const m = cleaned.match(/^(.+?)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)$/);
    if (!m) continue;
    const label = m[1].trim();
    if (!/[가-힣]/.test(label) || /예산구분|발행일|보조금|수익자/.test(label)) continue;
    rows.push({ index: i, page, label, compactLabel: norm(label), grant: toNumber(m[2]) * 1000, parent: toNumber(m[3]) * 1000, other: toNumber(m[4]) * 1000, total: toNumber(m[5]) * 1000, raw: line });
  }
  return rows;
}

function extractFormulaRows(lines, budgetRows) {
  const rows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].line;
    const m = line.match(/([0-9,]+)\s*원\s*\*\s*([^=]+?)\s*=\s*([0-9,]+)/);
    if (!m) continue;
    const nearLines = lines.slice(Math.max(0, i - 3), Math.min(lines.length, i + 2)).map(x => x.line);
    const windowText = nearLines.join(" ");
    const peopleMatches = [...line.matchAll(/\*\s*([0-9]+)\s*명/g)].map(x => Number(x[1]));
    const itemLabel = extractNearestItemLabel(nearLines);
    const section = nearestBudgetSection(i, budgetRows);
    rows.push({
      index: i,
      page: lines[i].page,
      line,
      itemLabel,
      section,
      windowText,
      compactWindow: norm(windowText),
      unit: toNumber(m[1]),
      total: toNumber(m[3]),
      people: peopleMatches.length ? Math.max(...peopleMatches) : 0
    });
  }
  return rows;
}

function extractNearestItemLabel(nearLines) {
  for (let i = nearLines.length - 2; i >= 0; i--) {
    let s = nearLines[i]
      .replace(/\(본예산\)/g, " ")
      .replace(/\(보조금및지원금\)|\(수익자부담금\)|\(그밖의수입\)/g, " ")
      .replace(/[()]/g, " ")
      .trim();
    s = s.replace(/^(교|환|운|경|직)\s+/, "").replace(/\s*(교|환|운|경|직)\s*$/, "").trim();
    if (!s || /^[0-9,]+/.test(s) || /예산구분|발행일|보조금|수익자/.test(s)) continue;
    if (/[가-힣]/.test(s) && !/[0-9,]+\s*원/.test(s)) return s.replace(/\s+/g, "");
  }
  return "";
}

function nearestBudgetSection(lineIndex, budgetRows) {
  const before = budgetRows.filter(r => r.index < lineIndex);
  const recent = before.slice(-8).map(r => r.label);
  const allowed = [...recent].reverse().find(x => RULES.allowedPersonnelBuckets.some(k => norm(x).includes(norm(k))));
  const leaf = recent[recent.length - 1] || "";
  return { path: recent, leaf, allowedBucket: allowed || "" };
}

function analyze(excel, pdf) {
  const results = [];
  const groups = buildExcelGroups(excel);

  comparePay(results, "교원급여", groups.teacherPay, getPdfFormulaAmount(pdf, ["교원급여"], ["방과후"]), getPdfPeople(pdf, ["교원급여"], ["방과후"]), true);
  comparePay(results, "방과후교원급여", groups.afterSchoolTeacherPay, getPdfFormulaAmount(pdf, ["방과후교원급여", "방과후급여"]), getPdfPeople(pdf, ["방과후교원급여", "방과후급여"]), true);
  comparePay(results, "교원급여 전체", groups.teacherPayTotal, getPdfItemAmount(pdf, ["교원급여"]), null, false);
  comparePay(results, "직원급여", groups.staffPay, getPdfItemAmount(pdf, ["직원급여"]), getPdfPeople(pdf, ["직원급여"], []), false);

  checkAllowances(results, groups.teacherAllowances, pdf, "교원수당", "교원");
  checkAllowances(results, groups.staffAllowances, pdf, "직원수당", "직원");
  checkRetirement(results, excel, pdf);
  checkMisclassifiedPersonnel(results, pdf);

  const debug = buildDebug(excel, groups, pdf);
  return { results, debug };
}

function buildExcelGroups(excelOrData) {
  // v7: 엑셀은 교직원보수일람표/보수기준 시트만 사용한다.
  // 사용자가 알려준 구조에 따라 "직명" 아래~소계(교원) 전까지 교원,
  // 그 아래~소계(직원/일반직) 전까지 직원을 읽고, 본봉/기본급은 연간액으로 본다.
  const salaryData = excelOrData?.salaryData || excelOrData;
  if (!salaryData || !salaryData.sections) {
    return buildExcelGroupsFromRows(Array.isArray(excelOrData) ? excelOrData : []);
  }

  const teacherRowsAll = salaryData.sections.teacher?.rows || [];
  const staffRows = salaryData.sections.staff?.rows || [];
  const regularTeacherRows = teacherRowsAll.filter(r => !norm(r.role).includes("방과후"));
  const afterSchoolRows = teacherRowsAll.filter(r => norm(r.role).includes("방과후"));

  const salaryTeacher = salaryData.sections.teacher || blankSection("교원");
  const salaryStaff = salaryData.sections.staff || blankSection("직원");

  return {
    teacherPay: {
      amount: sumBase(regularTeacherRows).amount,
      people: regularTeacherRows.length,
      rows: regularTeacherRows,
      source: "보수기준 시트 교원 구간 중 방과후 제외: 본봉/기본급 연간액 합계"
    },
    teacherPayTotal: {
      amount: sumBase(teacherRowsAll).amount,
      people: teacherRowsAll.length,
      rows: teacherRowsAll,
      source: "보수기준 시트 교원 구간 전체: 본봉/기본급 연간액 합계"
    },
    afterSchoolTeacherPay: {
      amount: sumBase(afterSchoolRows).amount,
      people: afterSchoolRows.length,
      rows: afterSchoolRows,
      source: "보수기준 시트 교원 구간 중 직명에 '방과후' 포함"
    },
    staffPay: {
      amount: sumBase(staffRows).amount,
      people: staffRows.length,
      rows: staffRows,
      source: "보수기준 시트 직원 구간: 본봉/기본급 연간액 합계"
    },
    teacherAllowances: salaryTeacher.subtotal?.allowances?.length ? salaryTeacher.subtotal.allowances : collectAllowances(teacherRowsAll),
    staffAllowances: salaryStaff.subtotal?.allowances?.length ? salaryStaff.subtotal.allowances : collectAllowances(staffRows),
    staffRoles: staffRows.map(r => r.role).filter(Boolean),
    parseInfo: salaryData
  };
}

function buildExcelGroupsFromRows(rows) {
  const regularTeachers = rows.filter(r => isTeacher(r.role) && !norm(r.role).includes("방과후"));
  const afterSchoolTeachers = rows.filter(r => norm(r.role).includes("방과후"));
  const staff = rows.filter(r => !regularTeachers.includes(r) && !afterSchoolTeachers.includes(r));
  return {
    teacherPay: sumBase(regularTeachers),
    afterSchoolTeacherPay: sumBase(afterSchoolTeachers),
    staffPay: sumBase(staff),
    teacherAllowances: collectAllowances(regularTeachers),
    staffAllowances: collectAllowances(staff),
    staffRoles: staff.map(r => r.role),
    parseInfo: null
  };
}

function isTeacher(role) {
  const r = norm(role);
  return r.includes("원장") || r.includes("원감") || r.includes("정교사") || r.includes("교사") || r.includes("교원") || r.includes("담임");
}

function sumBase(rows) {
  return { amount: rows.reduce((a, r) => a + (r.basePay || 0), 0), people: rows.filter(r => (r.basePay || 0) > 0).length, rows };
}

function collectAllowances(rows) {
  const map = new Map();
  rows.forEach(row => {
    Object.entries(row.allowances || {}).forEach(([name, value]) => {
      if (value <= 0 || isLegalBurden(name)) return;
      if (!hasAny(name, RULES.allowanceHeaders) && !RULES.directAllowanceAliases[name]) return;
      map.set(name, (map.get(name) || 0) + value);
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
  const rows = findPdfFormulaRows(pdf, aliases, excludeAliases);
  return rows.reduce((a, r) => a + r.total, 0);
}

function findPdfFormulaRows(pdf, aliases, excludeAliases = []) {
  const aliasNorms = aliases.map(norm);
  const excludeNorms = excludeAliases.map(norm);
  return pdf.formulaRows.filter(r => {
    const label = norm(r.itemLabel);
    const win = r.compactWindow;
    const matches = aliasNorms.some(a => label.includes(a) || win.includes(a));
    const excluded = excludeNorms.some(e => label.includes(e) || win.includes(e));
    return matches && !excluded;
  });
}

function getPdfPeople(pdf, aliases, excludeAliases = []) {
  const rows = findPdfFormulaRows(pdf, aliases, excludeAliases).filter(r => r.people > 0);
  if (!rows.length) return 0;
  // 같은 산출항목이 보조금/수익자부담금 등 재원별로 나뉘면 인원은 합산하지 않고 항목별 최대값을 사용한다.
  const byItem = new Map();
  rows.forEach(r => {
    const key = norm(r.itemLabel || aliases[0] || r.windowText);
    byItem.set(key, Math.max(byItem.get(key) || 0, r.people));
  });
  return [...byItem.values()].reduce((a, v) => a + v, 0);
}

function comparePay(results, label, excelGroup, pdfAmount, pdfPeople, checkPeople) {
  const amountDiff = excelGroup.amount - pdfAmount;
  const peopleDiff = (pdfPeople == null ? 0 : excelGroup.people - pdfPeople);
  let status = "ok", message = "금액 일치";
  if (excelGroup.amount === 0) {
    status = "warn"; message = "엑셀 금액 자동 추출 실패";
  } else if (pdfAmount === 0) {
    status = "bad"; message = "PDF 편성 금액 없음";
  } else if (!approxEqual(excelGroup.amount, pdfAmount)) {
    status = "bad"; message = `금액 차이 ${fmt(amountDiff)}원`;
  }
  if (checkPeople && pdfPeople != null && pdfPeople > 0 && peopleDiff !== 0) {
    status = "bad";
    message += `${message ? " / " : ""}인원수 불일치: PDF ${pdfPeople}명, 엑셀 ${excelGroup.people}명`;
  }
  results.push({ category: "보수", item: label, excelAmount: excelGroup.amount, pdfAmount, excelPeople: excelGroup.people, pdfPeople, status, message, evidence: excelGroup.source || "" });
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

  direct.forEach(a => results.push({ category: "수당", item: `${prefix} ${a.name}`, excelAmount: a.amount, pdfAmount: a.pdfAmount, status: "ok", message: "개별 편성 일치", evidence: "" }));

  const missingSum = missing.reduce((a, b) => a + b.amount, 0);
  if (missing.length === 0) return;
  const bucketAmount = getPdfFormulaAmount(pdf, [bucketName]);
  if (bucketAmount && approxEqual(bucketAmount, missingSum)) {
    results.push({ category: "수당", item: bucketName, excelAmount: missingSum, pdfAmount: bucketAmount, status: "info", message: `통합편성(${missing.map(x => `${x.name} ${fmt(x.amount)}원`).join(" + ")} = ${fmt(missingSum)}원)`, evidence: `${bucketName} 산출기초 합계 ${fmt(bucketAmount)}원` });
  } else {
    results.push({ category: "수당", item: bucketName, excelAmount: missingSum, pdfAmount: bucketAmount, status: "warn", message: `추가확인필요: PDF에서 개별 확인되지 않은 수당 ${missing.length}건의 합계와 ${bucketName} 산출기초 금액이 일치하지 않습니다.`, evidence: missing.map(x => `${x.name} ${fmt(x.amount)}원`).join(" + ") });
  }
}

function allowanceAliasesFor(name, prefix) {
  const base = RULES.directAllowanceAliases[name] || [name];
  return Array.from(new Set(base.flatMap(x => [x, `${prefix}${x}`, `${prefix}${name}`])));
}

function checkRetirement(results, excel, pdf) {
  const retirementSheets = excel.retirementSheets.filter(s => norm(s.name).includes("퇴직") || hasAny(s.text, ["퇴직", "적립", "충당"]));
  const amount = retirementSheets.reduce((sum, s) => sum + s.matrix.flat().reduce((a, c) => a + Math.max(0, toNumber(c)), 0), 0);
  const hasRetirementAmount = amount > 0;
  const found = hasAny(pdf.fullText, ["퇴직", "퇴직금", "퇴직적립", "퇴직급여충당"]);
  results.push({
    category: "퇴직금",
    item: "퇴직적립금 편성 여부",
    status: !hasRetirementAmount ? "info" : (found ? "ok" : "bad"),
    message: !hasRetirementAmount ? "엑셀 퇴직금 관련 금액을 찾지 못했습니다." : (found ? "엑셀 퇴직금 관련 금액 있음 / PDF 퇴직 관련 편성 확인" : "엑셀 퇴직금 관련 금액 있음 / PDF 퇴직 관련 편성 없음"),
    excelAmount: null,
    pdfAmount: null,
    evidence: `퇴직 관련 시트: ${retirementSheets.map(s => s.name).join(", ") || "없음"}`
  });
}

function checkMisclassifiedPersonnel(results, pdf) {
  const suspicious = [];
  pdf.formulaRows.forEach(r => {
    if (isLegalBurden(r.windowText)) return;
    const item = specificPersonnelLabel(r.itemLabel || r.windowText);
    if (!item) return;
    const pathText = r.section.path.join(" > ");
    const allowed = r.section.allowedBucket || hasAny(pathText, RULES.allowedPersonnelBuckets);
    if (allowed) return;
    const location = r.section.leaf || "위치 확인 필요";
    // 적립금/시설비 등에서 '급여'가 실제로 들어간 명확한 산출기초가 아닐 때는 지적하지 않는다.
    if (!norm(r.line).includes("급여") && !norm(r.itemLabel).includes("인건비")) return;
    suspicious.push({ item, location, amount: r.total, page: r.page, formula: r.line });
  });
  const unique = Array.from(new Map(suspicious.map(x => [`${x.item}-${x.location}-${x.amount}-${x.page}`, x])).values());
  unique.forEach(x => results.push({
    category: "오편성",
    item: x.item,
    status: "bad",
    message: `${x.item} ${fmt(x.amount)}원을 ${x.location}에 편성`,
    excelAmount: null,
    pdfAmount: x.amount,
    evidence: `PDF ${x.page}쪽: ${x.formula}`
  }));
}

function specificPersonnelLabel(text) {
  const compact = norm(text);
  if (isLegalBurden(compact)) return "";
  if (compact.includes("영양사")) return "영양사급여";
  if (compact.includes("조리직원") || compact.includes("조리원") || compact.includes("조리사")) return "조리직원급여";
  if (compact.includes("차량기사")) return "차량기사급여";
  if (compact.includes("보조교사")) return "보조교사급여";
  if (compact.includes("환경미화")) return "환경미화원급여";
  return "";
}

function isLegalBurden(text) { return hasAny(text, RULES.excludedLegalBurdenKeywords); }
function hasAny(text, keywords) { const n = norm(text); return keywords.some(k => n.includes(norm(k))); }

function buildDebug(excel, groups, pdf) {
  const info = excel.salaryData || groups.parseInfo;
  const teacher = info?.sections?.teacher || blankSection("교원");
  const staff = info?.sections?.staff || blankSection("직원");
  return [
    `선택된 보수 시트: ${excel.salarySheet?.name || "확인 불가"}` ,
    `월급여/비월정수당 등 기타 시트: 검토 제외`,
    `보수 대상 행: ${excel.salaryRows.length}건`,
    `교원 구간: ${teacher.startRow || "?"}~${teacher.endRow || "?"}행 / 소계행 ${teacher.subtotalRow || "?"}`,
    `직원 구간: ${staff.startRow || "?"}~${staff.endRow || "?"}행 / 소계행 ${staff.subtotalRow || "?"}`,
    `본봉열: ${info?.meta?.baseCol != null ? info.meta.baseCol + 1 : "?"} / 수당열: ${info?.meta?.allowanceStart != null ? info.meta.allowanceStart + 1 : "?"}~${info?.meta?.allowanceEnd != null ? info.meta.allowanceEnd + 1 : "?"}`,
    `퇴직 관련 시트 후보: ${excel.retirementSheets.map(s => s.name).join(", ") || "없음"}`,
    `PDF 예산 과목 행: ${pdf.budgetRows.length}건 / 산출기초 행: ${pdf.formulaRows.length}건`,
    "\n[엑셀 자동 집계]",
    `교원급여(방과후 제외): ${fmt(groups.teacherPay.amount)}원 / ${groups.teacherPay.people}명 / ${groups.teacherPay.source || ""}`,
    `교원급여 전체: ${fmt(groups.teacherPayTotal?.amount || 0)}원 / ${groups.teacherPayTotal?.people || 0}명`,
    `방과후교원급여: ${fmt(groups.afterSchoolTeacherPay.amount)}원 / ${groups.afterSchoolTeacherPay.people}명 / ${groups.afterSchoolTeacherPay.source || ""}`,
    `직원급여: ${fmt(groups.staffPay.amount)}원 / ${groups.staffPay.people}명 / ${groups.staffPay.source || ""}`,
    `교원수당 후보: ${groups.teacherAllowances.map(a => `${a.name}=${fmt(a.amount)}`).join(", ") || "없음"}`,
    `직원수당 후보: ${groups.staffAllowances.map(a => `${a.name}=${fmt(a.amount)}`).join(", ") || "없음"}`,
    "\n[PDF 예산 과목 일부]",
    pdf.budgetRows.slice(0, 45).map(r => `${r.label}: ${fmt(r.total)}원`).join("\n"),
    "\n[PDF 산출기초 일부]",
    pdf.formulaRows.slice(0, 60).map(r => `${r.itemLabel || "항목미상"} / ${r.line} / 위치: ${r.section.path.join(" > ")}`).join("\n")
  ];
}

function renderReport(report) {
  const counts = { ok: 0, info: 0, warn: 0, bad: 0 };
  report.results.forEach(r => counts[r.status]++);
  const summary = $("summary");
  summary.hidden = false;
  summary.innerHTML = `
    <div class="metric">정상<strong>${counts.ok}</strong></div>
    <div class="metric">참고/통합<strong>${counts.info}</strong></div>
    <div class="metric">추가확인<strong>${counts.warn}</strong></div>
    <div class="metric">지적<strong>${counts.bad}</strong></div>
  `;
  const badgeText = { ok: "정상", info: "참고", warn: "추가확인", bad: "지적" };
  $("results").innerHTML = `
    <div class="table-wrap"><table>
      <thead><tr><th>구분</th><th>항목</th><th>엑셀 금액</th><th>PDF 금액</th><th>엑셀 인원</th><th>PDF 인원</th><th>판정</th><th>내용</th><th>근거</th></tr></thead>
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
          <td>${escapeHtml(r.evidence || "")}</td>
        </tr>`).join("")}</tbody>
    </table></div>`;
  $("debugLog").textContent = report.debug.join("\n");
}

function escapeHtml(s) { return String(s ?? "").replace(/[&<>"]/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[m])); }
