/* global XLSX, pdfjsLib */

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const $ = (id) => document.getElementById(id);
const norm = (s) => String(s ?? "").replace(/\s+/g, "").trim();
const fmt = (n) => Number.isFinite(n) ? n.toLocaleString("ko-KR") : "-";

function toNumber(v) {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (v && typeof v === "object") {
    if (typeof v.v === "number") return v.v;
    if (typeof v.w === "string") return toNumber(v.w);
  }
  const cleaned = String(v ?? "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function hasAny(text, keywords) {
  const n = norm(text);
  return keywords.some(k => n.includes(norm(k)));
}

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
  const wb = XLSX.read(data, {
    type: "array",
    cellDates: false,
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    cellStyles: true
  });

  const sheetStates = Object.fromEntries((wb.Workbook?.Sheets || []).map(s => [s.name, s.Hidden || 0]));
  const sheets = wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const matrix = ws && ws["!ref"] ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: false }) : [];
    const text = matrix.flat().map(v => String(v ?? "")).join(" ");
    const hidden = sheetStates[name] === 1 || sheetStates[name] === 2;
    return { name, hidden, matrix, text };
  });

  const visible = sheets.filter(s => !s.hidden);
  const salarySheet = pickSalarySheet(visible) || pickSalarySheet(sheets) || visible[0] || sheets[0];
  const retirementSheets = visible.filter(s => hasAny(s.name + " " + s.text, ["퇴직"]));
  return { sheets, visible, salarySheet, retirementSheets };
}

function pickSalarySheet(sheets) {
  const scored = sheets.map(s => {
    const structure = findSalaryStructure(s.matrix);
    let score = 0;
    if (hasAny(s.name, ["보수", "급여", "봉급", "교직원", "일람"])) score += 10;
    if (hasAny(s.text, ["직명"])) score += 2;
    if (hasAny(s.text, ["본봉", "기본급"])) score += 2;
    if (hasAny(s.text, ["지급액계"])) score += 2;
    if (structure.headerRowIndex >= 0) score += 30;
    if (structure.selfDrivingCol >= 0) score += 10;
    return { sheet: s, score };
  }).sort((a, b) => b.score - a.score);
  return scored[0]?.score > 0 ? scored[0].sheet : null;
}

function findSalaryStructure(matrix) {
  let headerRowIndex = -1;
  for (let r = 0; r < matrix.length; r++) {
    const rowText = norm(matrix[r].join(" "));
    if (rowText.includes("직명") && (rowText.includes("본봉") || rowText.includes("기본급"))) {
      headerRowIndex = r;
      break;
    }
  }
  const row = headerRowIndex >= 0 ? matrix[headerRowIndex] : [];
  const next = headerRowIndex >= 0 ? matrix[headerRowIndex + 1] || [] : [];
  const maxCols = Math.max(row.length, next.length, ...matrix.map(r => r.length));
  const headers = Array.from({ length: maxCols }, (_, i) => ({
    index: i,
    name: String([row[i], next[i]].filter(Boolean).join(" ")).replace(/\s+/g, " ").trim()
  }));
  const baseCol = headers.find(h => hasAny(h.name, ["본봉", "기본급"]))?.index ?? -1;
  const payTotalCol = headers.find(h => norm(h.name).includes("지급액계"))?.index ?? -1;
  const selfDrivingCol = headers.find(h => {
    const n = norm(h.name);
    return n.includes("자가") && (n.includes("운전") || n.includes("보조금"));
  })?.index ?? -1;
  return { headerRowIndex, headers, baseCol, payTotalCol, selfDrivingCol };
}

function buildSalaryModel(sheet) {
  const matrix = sheet?.matrix || [];
  const structure = findSalaryStructure(matrix);
  const { headerRowIndex, headers, baseCol, payTotalCol, selfDrivingCol } = structure;
  const teacherSubtotal = matrix.findIndex(row => row.some(c => /소계\s*\(\s*교원\s*\)/.test(String(c))));
  const staffSubtotal = matrix.findIndex(row => row.some(c => /소계\s*\(\s*(직원|일반직)\s*\)/.test(String(c))));
  const dataStart = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;

  const teacherRows = teacherSubtotal >= 0 ? matrix.slice(dataStart, teacherSubtotal).filter(isSalaryDataRow) : [];
  const staffRows = teacherSubtotal >= 0 && staffSubtotal >= 0 ? matrix.slice(teacherSubtotal + 1, staffSubtotal).filter(isSalaryDataRow) : [];

  const teacherSelfDriving = selfDrivingCol >= 0 ? sumRowsByColumn(teacherRows, selfDrivingCol) : 0;
  const staffSelfDriving = selfDrivingCol >= 0 ? sumRowsByColumn(staffRows, selfDrivingCol) : 0;
  const selfDrivingAllowance = teacherSelfDriving + staffSelfDriving;

  return {
    sheetName: sheet?.name || "확인 불가",
    headerRowIndex,
    baseCol,
    payTotalCol,
    selfDrivingCol,
    teacherSubtotal,
    staffSubtotal,
    teacherRows,
    staffRows,
    teacherSelfDriving,
    staffSelfDriving,
    selfDrivingAllowance,
    headers
  };
}

function isSalaryDataRow(row) {
  const txt = row.map(v => String(v ?? "")).join(" ");
  if (!txt.trim()) return false;
  if (/소계|합계|작성요령|비고|구분/.test(txt)) return false;
  return row.some(v => toNumber(v) > 0);
}

function sumRowsByColumn(rows, col) {
  return rows.reduce((sum, row) => sum + toNumber(row[col]), 0);
}

function hasRetirementAmount(excel) {
  return excel.retirementSheets.some(s => {
    if (!hasAny(s.name + " " + s.text, ["퇴직"])) return false;
    return s.matrix.flat().some(v => toNumber(v) > 0);
  });
}

async function parsePdf(file) {
  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const simpleText = content.items.map(item => item.str).join("\n");
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
    pages.push({ page: i, lines, text: lines.join("\n"), simpleText });
  }
  const fullText = pages.map(p => p.text + "\n" + p.simpleText).join("\n");
  return { pages, fullText, compactText: norm(fullText), lines: fullText.split(/\n+/).map(s => s.trim()).filter(Boolean) };
}

function analyze(excel, pdf) {
  const results = [];
  const debug = [];
  const salary = buildSalaryModel(excel.salarySheet);

  debug.push(`보이는 시트: ${excel.visible.map(s => s.name).join(", ") || "없음"}`);
  debug.push(`선택된 보수/급여 시트: ${salary.sheetName}`);
  debug.push(`헤더 행: ${salary.headerRowIndex >= 0 ? salary.headerRowIndex + 1 : "못 찾음"}`);
  debug.push(`자가운전보조금 열: ${salary.selfDrivingCol >= 0 ? salary.selfDrivingCol + 1 : "못 찾음"}`);
  debug.push(`교원 구간 행 수: ${salary.teacherRows.length}, 직원 구간 행 수: ${salary.staffRows.length}`);
  debug.push(`자가운전보조금 합계: 교원 ${fmt(salary.teacherSelfDriving)}원 + 직원 ${fmt(salary.staffSelfDriving)}원 = ${fmt(salary.selfDrivingAllowance)}원`);
  debug.push(`퇴직 관련 시트: ${excel.retirementSheets.map(s => s.name).join(", ") || "없음"}`);
  debug.push(`PDF 차량기사급여/통학차량 탐지: ${detectVehicleDriverInTransport(pdf)?.amount ? fmt(detectVehicleDriverInTransport(pdf).amount) + "원" : "없음"}`);

  checkSelfDriving(results, salary, pdf);
  checkVehicleDriver(results, pdf);
  checkRetirement(results, excel, pdf);

  if (!results.length) {
    results.push({ category: "정상", item: "중점 검토", status: "ok", message: "현재 중점 검토 항목에서 지적사항이 발견되지 않았습니다.", excelAmount: null, pdfAmount: null });
  }

  return { results, debug };
}

function checkSelfDriving(results, salary, pdf) {
  if (salary.selfDrivingAllowance <= 0) return;
  if (!hasAny(pdf.fullText, ["자가운전", "자가 운전", "자가운전보조금"])) {
    results.push({
      category: "미편성",
      item: "자가운전보조금",
      excelAmount: salary.selfDrivingAllowance,
      pdfAmount: 0,
      status: "bad",
      message: `자가운전보조금 예산 미편성 (엑셀 ${fmt(salary.selfDrivingAllowance)}원)`
    });
  }
}

function detectVehicleDriverInTransport(pdf) {
  const compact = pdf.compactText;
  const start = compact.indexOf(norm("통학차량이용비"));
  if (start < 0) return null;
  const endCandidates = ["특별급식비간식비", "적립금", "시설설비비품비", "세출합계"].map(k => compact.indexOf(norm(k), start + 1)).filter(i => i > start);
  const end = endCandidates.length ? Math.min(...endCandidates) : Math.min(compact.length, start + 2000);
  const section = compact.slice(start, end);
  const direct = section.match(/차량기사급여.*?=([0-9,]+)원?/);
  if (direct) return { amount: toNumber(direct[1]), section: "통학차량이용비" };

  const lines = pdf.lines;
  const idx = lines.findIndex(l => norm(l).includes(norm("통학차량이용비")));
  if (idx >= 0) {
    const window = lines.slice(idx, idx + 80);
    const joined = norm(window.join(" "));
    const m = joined.match(/차량기사급여.*?=([0-9,]+)원?/);
    if (m) return { amount: toNumber(m[1]), section: "통학차량이용비" };
  }
  return null;
}

function checkVehicleDriver(results, pdf) {
  const detected = detectVehicleDriverInTransport(pdf);
  if (!detected?.amount) return;
  results.push({
    category: "오편성",
    item: "차량기사급여",
    excelAmount: null,
    pdfAmount: detected.amount,
    status: "bad",
    message: `차량기사급여 ${fmt(detected.amount)}원을 통학차량이용비의 산출내역에 편성`
  });
}

function checkRetirement(results, excel, pdf) {
  const excelHasRetirement = hasRetirementAmount(excel);
  const pdfHasRetirement = hasAny(pdf.fullText, ["퇴직금", "퇴직적립", "퇴직 적립", "퇴직급여", "퇴직 급여", "퇴직충당", "퇴직 충당"]);
  if (excelHasRetirement && !pdfHasRetirement) {
    results.push({ category: "미편성", item: "퇴직적립금", status: "bad", message: "퇴직적립금 미편성", excelAmount: null, pdfAmount: 0 });
  } else if (excelHasRetirement && pdfHasRetirement) {
    results.push({ category: "정상", item: "퇴직적립금", status: "ok", message: "퇴직 관련 편성 항목 확인", excelAmount: null, pdfAmount: null });
  } else {
    results.push({ category: "추가확인", item: "퇴직적립금", status: "warn", message: "퇴직 관련 시트의 적립금액을 자동 확인하지 못했습니다.", excelAmount: null, pdfAmount: null });
  }
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
