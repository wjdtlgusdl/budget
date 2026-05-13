import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs";
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs";

const $ = (id) => document.getElementById(id);
const state = { rows: [], debug: null };

const CONFIG = {
  thousandUnitTolerance: 0,
  employeeGroups: [
    { group: "교원", aliases: ["원장", "원감", "교사", "정교사", "보육교사", "담임", "교원"], expectedPdf: ["교원급여", "교원인건비", "보육교직원인건비"] },
    { group: "방과후교사", aliases: ["방과후", "방과후교사", "방과후교원"], expectedPdf: ["방과후교원급여", "방과후교사급여", "방과후과정"] },
    { group: "직원", aliases: ["직원", "사무", "행정", "조리", "운전", "관리", "보조", "간호", "영양사"], expectedPdf: ["직원급여", "직원인건비", "기타직원인건비"] },
    { group: "영양사", aliases: ["영양사"], expectedPdf: ["영양사", "직원인건비", "급식비"] }
  ],
  payItems: [
    { key: "base", name: "본봉/기본급", aliases: ["본봉", "기본급", "급여"], compareAmount: true },
    { key: "overtime", name: "시간외수당", aliases: ["시간외", "초과근무"], compareAmount: true, allowance: true },
    { key: "research", name: "연구활동비", aliases: ["연구활동", "연구수당"], compareAmount: true, allowance: true },
    { key: "meal", name: "식대", aliases: ["식대", "급식비", "정액급식"], compareAmount: true, allowance: true },
    { key: "holiday", name: "명절휴가비", aliases: ["명절", "명절휴가"], compareAmount: true, allowance: true },
    { key: "teacherDay", name: "스승의날상여금", aliases: ["스승", "스승의날"], compareAmount: true, allowance: true },
    { key: "vacation", name: "방학휴가비", aliases: ["방학", "휴가비"], compareAmount: true, allowance: true },
    { key: "performance", name: "성과상여금", aliases: ["성과", "성과상여"], compareAmount: true, allowance: true }
  ],
  retirementSheetPatterns: ["퇴직", "퇴직금", "퇴직적립", "퇴직급여"],
  retirementPdfPatterns: ["퇴직금", "퇴직적립금", "퇴직급여", "퇴직급여충당", "퇴직충당금"],
  integratedAllowancePatterns: ["교원수당", "직원수당", "제수당", "각종수당", "처우개선", "수당"]
};

$("analyzeBtn").addEventListener("click", analyze);
$("downloadCsvBtn").addEventListener("click", downloadCsv);
$("downloadJsonBtn").addEventListener("click", downloadJson);
$("resetBtn").addEventListener("click", () => location.reload());

async function analyze() {
  const excelFile = $("excelFile").files[0];
  const pdfFile = $("pdfFile").files[0];
  if (!excelFile || !pdfFile) return setStatus("엑셀 파일과 PDF 파일을 모두 업로드하세요.", true);

  setStatus("엑셀과 PDF를 분석하는 중입니다...");
  try {
    const excel = await parseExcel(excelFile);
    const pdf = await parsePdf(pdfFile);
    const rows = buildReviewRows(excel, pdf);
    state.rows = rows;
    state.debug = { excel, pdfPreview: pdf.text.slice(0, 8000), pdfSections: pdf.sections.slice(0, 80) };
    renderResults(rows, state.debug);
    setStatus(`검토 완료: ${rows.length}개 항목`);
  } catch (err) {
    console.error(err);
    setStatus(`분석 중 오류가 발생했습니다: ${err.message}`, true);
  }
}

async function parseExcel(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: false, cellNF: false, cellStyles: false });
  const sheetMeta = workbook.Workbook?.Sheets || [];
  const visibleSheets = workbook.SheetNames.filter((name, idx) => !sheetMeta[idx]?.Hidden);
  const hiddenSheets = workbook.SheetNames.filter((name, idx) => !!sheetMeta[idx]?.Hidden);
  const sheets = visibleSheets.map((name) => {
    const ws = workbook.Sheets[name];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
    return { name, matrix, flatText: matrix.flat().join(" ") };
  });

  const records = [];
  const sheetTotals = [];
  for (const sheet of sheets) {
    const normalizedRows = normalizeExcelRows(sheet.matrix);
    for (const row of normalizedRows) {
      const text = row.join(" ");
      const numbers = row.map(parseMoney).filter((n) => n > 0);
      if (!numbers.length) continue;
      const group = detectGroup(text);
      const item = detectPayItem(text, sheet.name);
      const headcount = detectHeadcount(text) || 1;
      const amount = guessRowAmount(numbers);
      if (item || group !== "기타") {
        records.push({ sheet: sheet.name, group, itemKey: item?.key || "unknown", itemName: item?.name || "미분류", text, amount, headcount });
      }
    }
    sheetTotals.push({ sheet: sheet.name, amount: sumAllMoney(sheet.matrix), retirementLike: hasAny(sheet.name, CONFIG.retirementSheetPatterns) || hasAny(sheet.flatText, CONFIG.retirementSheetPatterns) });
  }

  const retirementAmount = sheetTotals.filter((s) => s.retirementLike).reduce((a, b) => a + b.amount, 0);
  const grouped = aggregateRecords(records);
  return { visibleSheets, hiddenSheets, records, grouped, retirementAmount, sheetTotals };
}

function normalizeExcelRows(matrix) {
  const rows = [];
  let lastLabel = "";
  for (const raw of matrix) {
    const row = raw.map((v) => String(v ?? "").trim());
    const text = row.join(" ").trim();
    if (!text) continue;
    const firstText = row.find((v) => v && !looksMoney(v));
    if (firstText && firstText.length > 1) lastLabel = firstText;
    rows.push([lastLabel, ...row]);
  }
  return rows;
}

function aggregateRecords(records) {
  const map = new Map();
  for (const r of records) {
    if (r.itemKey === "unknown") continue;
    const key = `${r.group}|${r.itemKey}`;
    if (!map.has(key)) map.set(key, { group: r.group, itemKey: r.itemKey, itemName: r.itemName, amount: 0, headcount: 0, sources: [] });
    const target = map.get(key);
    target.amount += r.amount;
    target.headcount += r.headcount;
    target.sources.push({ sheet: r.sheet, amount: r.amount, text: r.text.slice(0, 180) });
  }
  return Array.from(map.values()).filter((x) => x.amount > 0);
}

async function parsePdf(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const items = content.items.map((item) => item.str).filter(Boolean);
    pages.push({ page: p, text: items.join(" ") });
  }
  const text = pages.map((p) => `\n[PAGE ${p.page}]\n${p.text}`).join("\n");
  return { pages, text, sections: splitPdfSections(text) };
}

function splitPdfSections(text) {
  const lines = text.split(/\n|(?=\d+\.\s*)/).map((s) => s.trim()).filter(Boolean);
  return lines.map((line, idx) => ({ index: idx, text: line, amount: extractLastAmount(line), headcount: detectHeadcount(line), calcAmounts: extractCalculationAmounts(line) }));
}

function buildReviewRows(excel, pdf) {
  const rows = [];
  for (const g of excel.grouped) {
    const itemConfig = CONFIG.payItems.find((p) => p.key === g.itemKey);
    const expectedTerms = [...new Set([g.itemName, ...(itemConfig?.aliases || []), ...expectedTermsForGroup(g.group)])];
    const expectedHit = findPdfHits(pdf, expectedTerms, expectedTermsForGroup(g.group));
    const globalHit = expectedHit.length ? expectedHit : findPdfHits(pdf, expectedTerms, []);
    const hit = chooseBestHit(globalHit, g);

    const isIntegrated = itemConfig?.allowance && !hit && hasIntegratedAllowance(pdf.text, g.group);
    const pdfAmount = hit ? normalizePdfAmount(hit.amount || sumCalcAmounts(hit.calcAmounts)) : null;
    const diff = pdfAmount == null ? null : pdfAmount - g.amount;
    const moneyStatus = determineMoneyStatus(g.amount, pdfAmount, isIntegrated);
    const formation = determineFormation(hit, globalHit, isIntegrated, g, pdf);
    const final = determineFinal(moneyStatus, formation);
    const note = makeNote(g, hit, globalHit, isIntegrated, pdf);

    rows.push({
      group: g.group,
      item: g.itemName,
      excelAmount: g.amount,
      pdfAmount,
      diff,
      moneyStatus,
      formation,
      final,
      note
    });
  }

  if (excel.retirementAmount > 0) {
    const hit = findFirstPattern(pdf.text, CONFIG.retirementPdfPatterns);
    rows.push({
      group: "퇴직금",
      item: "퇴직금/퇴직적립금",
      excelAmount: excel.retirementAmount,
      pdfAmount: null,
      diff: null,
      moneyStatus: "편성 여부만 확인",
      formation: hit ? "편성되어 있음" : "미편성",
      final: hit ? "적정" : "부적정",
      note: hit ? `엑셀 시트에 퇴직금액 존재, PDF에서 '${hit}' 확인. 금액 비교 제외.` : "엑셀 시트에 퇴직금액 존재, 퇴직금 예산 미편성, 부적정"
    });
  }

  return rows.sort((a, b) => scoreFinal(b.final) - scoreFinal(a.final));
}

function expectedTermsForGroup(group) {
  const cfg = CONFIG.employeeGroups.find((x) => x.group === group);
  return cfg?.expectedPdf || [];
}

function findPdfHits(pdf, itemTerms, sectionTerms) {
  const terms = [...itemTerms, ...sectionTerms].filter(Boolean);
  const hits = [];
  for (const sec of pdf.sections) {
    const text = normalize(sec.text);
    const itemMatched = itemTerms.some((t) => text.includes(normalize(t)));
    const sectionMatched = sectionTerms.length === 0 || sectionTerms.some((t) => text.includes(normalize(t)));
    if (itemMatched && sectionMatched) hits.push(sec);
  }
  return hits;
}

function chooseBestHit(hits, g) {
  if (!hits.length) return null;
  const expectedK = Math.floor(g.amount / 1000);
  return hits
    .map((h) => {
      const amount = normalizePdfAmount(h.amount || sumCalcAmounts(h.calcAmounts));
      const amountScore = amount == null ? 999999999 : Math.abs(Math.floor(amount / 1000) - expectedK);
      const headScore = h.headcount ? Math.abs(h.headcount - g.headcount) : 99;
      return { h, score: amountScore + headScore * 1000 };
    })
    .sort((a, b) => a.score - b.score)[0].h;
}

function normalizePdfAmount(amount) {
  if (amount == null || Number.isNaN(amount)) return null;
  // 예산서 합계가 천원 단위로 표시된 경우 원 단위로 환산한다.
  if (amount > 0 && amount < 10000000) return amount * 1000;
  return amount;
}

function determineMoneyStatus(excelAmount, pdfAmount, integrated) {
  if (integrated) return "비교불가";
  if (pdfAmount == null) return "미편성";
  const excelK = Math.floor(excelAmount / 1000);
  const pdfK = Math.floor(pdfAmount / 1000);
  if (Math.abs(excelK - pdfK) <= CONFIG.thousandUnitTolerance) return "일치";
  return pdfK < excelK ? "과소편성" : "과다편성";
}

function determineFormation(hit, hits, integrated, g, pdf) {
  if (integrated) return "통합 편성 의심";
  if (!hit) return findElsewhere(pdf.text, g) || "미편성";
  const matchingCalcHits = hits.filter((h) => normalizePdfAmount(h.amount || sumCalcAmounts(h.calcAmounts)) != null);
  const amountSum = matchingCalcHits.reduce((a, h) => a + (normalizePdfAmount(h.amount || sumCalcAmounts(h.calcAmounts)) || 0), 0);
  const headSum = matchingCalcHits.reduce((a, h) => a + (h.headcount || 0), 0);
  if (matchingCalcHits.length >= 2 && Math.floor(amountSum / 1000) === Math.floor(g.amount / 1000) && (!g.headcount || !headSum || headSum === g.headcount)) {
    return "분리 편성 되어 있음";
  }
  const elsewhere = findElsewhere(pdf.text, g);
  if (elsewhere && !hits.some((h) => expectedTermsForGroup(g.group).some((t) => normalize(h.text).includes(normalize(t))))) return elsewhere;
  return "단일 편성";
}

function determineFinal(moneyStatus, formation) {
  if (formation === "미편성" || moneyStatus === "미편성" || moneyStatus === "과소편성" || moneyStatus === "과다편성") return "부적정";
  if (formation.includes("의심") || formation.includes("타 목") || formation.includes("급식비") || moneyStatus === "비교불가") return "확인 필요";
  return "적정";
}

function makeNote(g, hit, hits, integrated, pdf) {
  if (integrated) return `${g.itemName}이 개별 항목이 아니라 교원수당/직원수당 등으로 통합 편성되었을 가능성이 있습니다.`;
  if (!hit) return "예상 목 및 PDF 전체에서 명확한 산출내역을 찾지 못했습니다.";
  const pieces = [];
  if (hits.length >= 2) pieces.push(`관련 산출항목 ${hits.length}건 확인`);
  if (hit.headcount) pieces.push(`산출기초 인원 ${hit.headcount}명 확인`);
  pieces.push(hit.text.slice(0, 220));
  return pieces.join(" · ");
}

function findElsewhere(text, g) {
  const terms = [g.itemName, ...(CONFIG.payItems.find((p) => p.key === g.itemKey)?.aliases || [])];
  const normText = normalize(text);
  if (!terms.some((t) => normText.includes(normalize(t)))) return null;
  const foodTerms = ["급식비", "급식운영", "학교급식", "영양사"];
  if (foodTerms.some((t) => normText.includes(normalize(t))) && g.group === "영양사") return "급식비에 편성";
  return "타 목 편성 가능성";
}

function hasIntegratedAllowance(text, group) {
  const norm = normalize(text);
  const groupTerms = expectedTermsForGroup(group).map(normalize);
  return CONFIG.integratedAllowancePatterns.some((p) => norm.includes(normalize(p))) || groupTerms.some((t) => norm.includes(t + normalize("수당")));
}

function findFirstPattern(text, patterns) {
  const norm = normalize(text);
  return patterns.find((p) => norm.includes(normalize(p))) || null;
}

function detectGroup(text) {
  const norm = normalize(text);
  const sorted = [...CONFIG.employeeGroups].sort((a, b) => b.aliases.join(" ").length - a.aliases.join(" ").length);
  return sorted.find((g) => g.aliases.some((a) => norm.includes(normalize(a))))?.group || "기타";
}

function detectPayItem(text, sheetName = "") {
  const norm = normalize(`${sheetName} ${text}`);
  return CONFIG.payItems.find((p) => p.aliases.some((a) => norm.includes(normalize(a)))) || null;
}

function detectHeadcount(text) {
  const m = String(text).match(/(\d{1,3})\s*(명|인)/);
  return m ? Number(m[1]) : 0;
}

function extractLastAmount(text) {
  const amounts = extractMoneyValues(text);
  return amounts.length ? amounts[amounts.length - 1] : null;
}

function extractCalculationAmounts(text) {
  const amounts = [];
  const calcMatches = String(text).matchAll(/=\s*([\d,]+)\s*원?/g);
  for (const m of calcMatches) amounts.push(parseMoney(m[1]));
  return amounts.filter((n) => n > 0);
}

function sumCalcAmounts(amounts) {
  return amounts?.reduce((a, b) => a + b, 0) || null;
}

function sumAllMoney(matrix) {
  return matrix.flat().map(parseMoney).filter((n) => n > 0).reduce((a, b) => a + b, 0);
}

function extractMoneyValues(text) {
  const matches = String(text).match(/[\d,]{2,}\s*(원|천원)?/g) || [];
  return matches.map(parseMoney).filter((n) => n > 0);
}

function parseMoney(value) {
  if (value == null) return 0;
  const s = String(value).replace(/\s/g, "");
  if (!/[0-9]/.test(s)) return 0;
  const n = Number(s.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? Math.abs(n) : 0;
}

function looksMoney(v) {
  return /^[\s\d,.-]+(원|천원)?\s*$/.test(String(v));
}

function guessRowAmount(numbers) {
  return Math.max(...numbers);
}

function hasAny(text, patterns) {
  const norm = normalize(text);
  return patterns.some((p) => norm.includes(normalize(p)));
}

function normalize(s) {
  return String(s || "").replace(/\s+/g, "").toLowerCase();
}

function renderResults(rows, debug) {
  const tbody = $("resultsTable").querySelector("tbody");
  tbody.innerHTML = "";
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.group)}</td>
      <td>${escapeHtml(r.item)}</td>
      <td class="amount">${formatWon(r.excelAmount)}</td>
      <td class="amount">${r.pdfAmount == null ? "-" : formatWon(r.pdfAmount)}</td>
      <td class="amount">${r.diff == null ? "-" : formatWon(r.diff)}</td>
      <td>${badge(r.moneyStatus)}</td>
      <td>${escapeHtml(r.formation)}</td>
      <td>${badge(r.final)}</td>
      <td>${escapeHtml(r.note)}</td>
    `;
    tbody.appendChild(tr);
  }
  renderSummary(rows, debug);
  $("resultsCard").hidden = false;
  $("summaryCard").hidden = false;
  $("debugCard").hidden = false;
  $("debugOutput").textContent = JSON.stringify(debug, null, 2);
  $("downloadCsvBtn").disabled = false;
  $("downloadJsonBtn").disabled = false;
}

function renderSummary(rows, debug) {
  const counts = rows.reduce((acc, r) => { acc[r.final] = (acc[r.final] || 0) + 1; return acc; }, {});
  const html = [
    metric("전체 항목", rows.length),
    metric("적정", counts["적정"] || 0),
    metric("확인 필요", counts["확인 필요"] || 0),
    metric("부적정", counts["부적정"] || 0),
    metric("제외된 숨김 시트", debug.excel.hiddenSheets.length)
  ].join("");
  $("summary").innerHTML = html;
}

function metric(label, value) { return `<div class="metric"><strong>${value}</strong><span>${label}</span></div>`; }
function badge(text) {
  const cls = text === "적정" || text === "일치" || text === "편성 여부만 확인" ? "ok" : text === "부적정" || text.includes("과") || text === "미편성" ? "bad" : text.includes("확인") || text.includes("비교") ? "warn" : "info";
  return `<span class="badge ${cls}">${escapeHtml(text)}</span>`;
}
function formatWon(n) { return Number(n || 0).toLocaleString("ko-KR") + "원"; }
function escapeHtml(s) { return String(s ?? "").replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c])); }
function setStatus(msg, isError = false) { const el = $("status"); el.textContent = msg; el.style.color = isError ? "#b91c1c" : "#667085"; }
function scoreFinal(final) { return final === "부적정" ? 3 : final === "확인 필요" ? 2 : 1; }

function downloadCsv() {
  const headers = ["그룹", "항목", "엑셀 기준", "PDF 확인", "차이", "금액검증", "편성형태", "최종판정", "비고"];
  const lines = [headers, ...state.rows.map((r) => [r.group, r.item, r.excelAmount, r.pdfAmount ?? "", r.diff ?? "", r.moneyStatus, r.formation, r.final, r.note])]
    .map((row) => row.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","));
  downloadBlob("budget-review-result.csv", "text/csv;charset=utf-8", "\ufeff" + lines.join("\n"));
}
function downloadJson() { downloadBlob("budget-review-result.json", "application/json", JSON.stringify({ rows: state.rows, debug: state.debug }, null, 2)); }
function downloadBlob(filename, type, content) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
