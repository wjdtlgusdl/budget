const $ = (id) => document.getElementById(id);
const state = { rows: [], debug: null };

const CONFIG = {
  teacherRole: /원장|원감|교사|정교사|담임|보육교사/i,
  afterRole: /방과후/i,
  staffRole: /직원|사무|행정|조리|운전|관리|보조|간호|영양사/i,
  retirementSheetPatterns: ["퇴직", "퇴직금", "퇴직적립", "퇴직급여"],
  retirementPdfPatterns: ["퇴직금", "퇴직적립금", "퇴직급여", "퇴직급여충당", "퇴직충당금"],
  pdfSectionTerms: {
    teacherBase: ["교원급여", "교원인건비", "교원기본급"],
    afterBase: ["방과후교원급여", "방과후교사급여", "방과후과정"],
    staffBase: ["직원급여", "직원인건비", "직원기본급"],
    nutritionist: ["영양사", "급식비", "급식운영"],
    teacherAllowance: ["교원수당", "교원정액급식비", "교원명절휴가비", "교원스승의날상여금", "교원방학휴가비", "교원성과상여금"],
    staffAllowance: ["직원수당", "직원정액급식비", "직원명절휴가비", "직원성과상여금"]
  },
  allowanceColumns: [
    { key: "overtime", name: "시간외수당", aliases: ["시간외"] },
    { key: "research", name: "연구활동비", aliases: ["연구활동"] },
    { key: "meal", name: "식대", aliases: ["식대", "정액급식"] },
    { key: "holiday", name: "명절휴가비", aliases: ["명절"] },
    { key: "teacherDay", name: "스승의날상여금", aliases: ["스승"] },
    { key: "vacation", name: "방학휴가비", aliases: ["방학", "휴가"] },
    { key: "performance", name: "성과상여금", aliases: ["성과"] }
  ]
};

document.addEventListener("DOMContentLoaded", init);

function init() {
  try {
    if (window.pdfjsLib) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
    }
    const analyzeBtn = $("analyzeBtn");
    if (!analyzeBtn) throw new Error("실행 버튼을 찾지 못했습니다. index.html과 app.js가 같은 폴더에 있는지 확인하세요.");
    analyzeBtn.addEventListener("click", analyze);
    $("downloadCsvBtn").addEventListener("click", downloadCsv);
    $("downloadJsonBtn").addEventListener("click", downloadJson);
    $("resetBtn").addEventListener("click", () => location.reload());
    setStatus("파일을 업로드한 뒤 검토를 실행하세요.");
  } catch (err) {
    const el = document.getElementById("status");
    if (el) {
      el.textContent = `초기화 오류: ${err.message}`;
      el.style.color = "#b91c1c";
    }
    console.error(err);
  }
}

async function analyze() {
  const excelFile = $("excelFile").files[0];
  const pdfFile = $("pdfFile").files[0];
  if (!excelFile || !pdfFile) return setStatus("엑셀 파일과 PDF 파일을 모두 업로드하세요.", true);
  setStatus("분석 중입니다...");
  try {
    const excel = await parseExcel(excelFile);
    const pdf = await parsePdf(pdfFile);
    const rows = buildReviewRows(excel, pdf);
    state.rows = rows;
    state.debug = {
      visibleSheets: excel.visibleSheets,
      hiddenSheets: excel.hiddenSheets,
      extractedExcelGroups: excel.groups,
      retirementAmountDetected: excel.retirementAmount,
      pdfPreview: pdf.text.slice(0, 10000),
      pdfLinesPreview: pdf.lines.slice(0, 120)
    };
    renderResults(rows, state.debug);
    setStatus(`검토 완료: ${rows.length}개 항목`);
  } catch (err) {
    console.error(err);
    setStatus(`분석 중 오류가 발생했습니다: ${err.message}`, true);
  }
}

async function parseExcel(file) {
  if (!window.XLSX) throw new Error("XLSX 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 차단 여부를 확인하세요.");
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", raw: false, cellDates: false });
  const sheetMeta = workbook.Workbook?.Sheets || [];
  const visibleSheets = workbook.SheetNames.filter((name, idx) => !sheetMeta[idx]?.Hidden);
  const hiddenSheets = workbook.SheetNames.filter((name, idx) => !!sheetMeta[idx]?.Hidden);
  const sheets = visibleSheets.map((name) => ({
    name,
    matrix: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: false, defval: "" })
  }));

  const peopleRows = [];
  const retirementSheets = [];
  for (const sheet of sheets) {
    const flat = sheet.matrix.flat().join(" ");
    if (hasAny(`${sheet.name} ${flat}`, CONFIG.retirementSheetPatterns)) {
      const amount = sumVisibleMoney(sheet.matrix);
      if (amount > 0) retirementSheets.push({ sheet: sheet.name, amount });
    }
    peopleRows.push(...extractPeopleRows(sheet));
  }

  const groups = aggregatePeople(peopleRows);
  return {
    visibleSheets,
    hiddenSheets,
    peopleRows,
    groups,
    retirementAmount: retirementSheets.reduce((a, b) => a + b.amount, 0),
    retirementSheets
  };
}

function extractPeopleRows(sheet) {
  const rows = sheet.matrix.map((r) => r.map((v) => String(v ?? "").trim()));
  const headerIdx = rows.findIndex((r) => r.join(" ").includes("직명") && r.join(" ").includes("성") && (r.join(" ").includes("본봉") || r.join(" ").includes("기본급")));
  if (headerIdx < 0) return [];

  const headerTextByCol = {};
  for (let c = 0; c < Math.max(...rows.map((r) => r.length)); c++) {
    headerTextByCol[c] = rows.slice(headerIdx, Math.min(headerIdx + 3, rows.length)).map((r) => r[c] || "").join(" ");
  }
  const col = {
    role: findCol(headerTextByCol, ["직명"]),
    name: findCol(headerTextByCol, ["성 명", "성명"]),
    month: findCol(headerTextByCol, ["근무", "월수"]),
    base: findCol(headerTextByCol, ["본봉", "기본급"]),
    overtime: findCol(headerTextByCol, ["시간외"]),
    research: findCol(headerTextByCol, ["연구활동"]),
    meal: findCol(headerTextByCol, ["식대"]),
    holiday: findCol(headerTextByCol, ["명절"]),
    teacherDay: findCol(headerTextByCol, ["스승"]),
    vacation: findCol(headerTextByCol, ["방학", "휴가"]),
    performance: findCol(headerTextByCol, ["성과"])
  };
  if (col.role < 0 || col.base < 0) return [];

  const out = [];
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const role = row[col.role] || "";
    if (!role || /소계|합계|계\b/.test(role)) continue;
    const base = money(row[col.base]);
    if (!base) continue;
    const month = Math.max(1, Math.min(12, money(row[col.month]) || inferMonth(base)));
    const group = classifyGroup(role);
    const annualBase = base < 10000000 ? base * month : base;
    const record = {
      sheet: sheet.name,
      role,
      name: col.name >= 0 ? row[col.name] : "",
      group,
      month,
      base: annualBase,
      allowances: {}
    };
    for (const a of CONFIG.allowanceColumns) {
      const v = col[a.key] >= 0 ? money(row[col[a.key]]) : 0;
      record.allowances[a.key] = v < 10000000 ? v * month : v;
    }
    out.push(record);
  }
  return out;
}

function findCol(headerTextByCol, terms) {
  const entries = Object.entries(headerTextByCol);
  const termList = terms.map(normalize);
  const hit = entries.find(([, text]) => termList.every((t) => normalize(text).includes(t))) ||
              entries.find(([, text]) => termList.some((t) => normalize(text).includes(t)));
  return hit ? Number(hit[0]) : -1;
}

function classifyGroup(role) {
  if (CONFIG.afterRole.test(role)) return "방과후교사";
  if (/영양사/i.test(role)) return "영양사";
  if (CONFIG.teacherRole.test(role)) return "교원";
  if (CONFIG.staffRole.test(role)) return "직원";
  return "기타";
}

function aggregatePeople(peopleRows) {
  const make = (group, item, key) => ({ group, item, key, amount: 0, headcount: 0, sources: [] });
  const map = new Map();
  function add(group, item, key, amount, person) {
    if (!amount || amount <= 0 || group === "기타") return;
    const k = `${group}|${key}`;
    if (!map.has(k)) map.set(k, make(group, item, key));
    const g = map.get(k);
    g.amount += amount;
    g.headcount += 1;
    g.sources.push(`${person.role} ${person.name} ${formatWon(amount)}`);
  }
  for (const p of peopleRows) {
    add(p.group, "본봉/기본급", "base", p.base, p);
    for (const a of CONFIG.allowanceColumns) add(p.group, a.name, a.key, p.allowances[a.key], p);
  }
  return Array.from(map.values()).filter((g) => g.amount > 0);
}

async function parsePdf(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  if (!window.pdfjsLib) throw new Error("PDF.js 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 차단 여부를 확인하세요.");
  const pdf = await window.pdfjsLib.getDocument({ data }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const sorted = content.items
      .map((it) => ({ text: it.str, x: it.transform[4], y: Math.round(it.transform[5]) }))
      .filter((it) => it.text && it.text.trim())
      .sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    let curY = null, cur = [];
    for (const it of sorted) {
      if (curY === null || Math.abs(it.y - curY) <= 2) {
        curY = curY ?? it.y;
        cur.push(it.text);
      } else {
        lines.push(cur.join(" "));
        curY = it.y; cur = [it.text];
      }
    }
    if (cur.length) lines.push(cur.join(" "));
    pages.push({ page: p, lines, text: lines.join("\n") });
  }
  const lines = pages.flatMap((p) => p.lines.map((line) => `[${p.page}] ${line}`));
  return { pages, lines, text: lines.join("\n") };
}

function buildReviewRows(excel, pdf) {
  const rows = [];
  for (const g of excel.groups) {
    if (g.key === "base") rows.push(reviewBase(g, pdf));
    else rows.push(reviewAllowance(g, pdf));
  }
  if (excel.retirementAmount > 0) {
    const hit = findLine(pdf.lines, CONFIG.retirementPdfPatterns);
    rows.push({
      group: "퇴직금", item: "퇴직금/퇴직적립금", excelAmount: excel.retirementAmount,
      pdfAmount: null, diff: null, moneyStatus: "편성 여부만 확인",
      formation: hit ? "편성되어 있음" : "미편성", final: hit ? "적정" : "부적정",
      note: hit ? `엑셀 시트에 퇴직금액 존재, PDF에서 확인: ${clip(hit)}` : "엑셀 시트에 퇴직금액 존재, 퇴직금 예산 미편성, 부적정"
    });
  }
  return rows.sort((a, b) => scoreFinal(b.final) - scoreFinal(a.final));
}

function reviewBase(g, pdf) {
  const terms = baseTerms(g.group);
  const hits = findRelevantBlocks(pdf.lines, terms);
  const calcHits = hits.map((h) => ({ line: h, amounts: extractCalcResults(h), headcount: extractHeadcount(h) }))
    .filter((h) => h.amounts.length || extractAmounts(h.line).length);
  const allAmounts = calcHits.flatMap((h) => h.amounts.length ? h.amounts : extractAmounts(h.line));
  const exactAmounts = allAmounts.filter((a) => sameThousand(a, g.amount));
  const sum = allAmounts.reduce((a, b) => a + normalizePdfAmount(b), 0);
  let pdfAmount = exactAmounts[0] ? normalizePdfAmount(exactAmounts[0]) : (sameThousand(sum, g.amount) ? sum : bestAmount(allAmounts, g.amount));
  let moneyStatus = amountStatus(g.amount, pdfAmount);
  let formation = "미편성";
  let note = "예상 목과 PDF 전체에서 명확한 산출내역을 찾지 못했습니다.";

  const headSum = calcHits.reduce((a, h) => a + (h.headcount || 0), 0);
  if (pdfAmount != null) {
    formation = calcHits.length >= 2 && sameThousand(sum, g.amount) ? "분리 편성 되어 있음" : "단일 편성";
    note = `${calcHits.length || hits.length}건 확인${headSum ? `, 산출기초 인원 합계 ${headSum}명` : ""}. ${clip((hits[0] || ""))}`;
  } else if (g.group === "영양사") {
    const food = findLine(pdf.lines, ["영양사", "급식비", "급식운영"]);
    if (food) { formation = "급식비에 편성"; moneyStatus = "비교불가"; note = `직원 인건비에서 특정하기 어렵지만 PDF 전체에서 확인: ${clip(food)}`; }
  }
  return finish(g, pdfAmount, moneyStatus, formation, note);
}

function reviewAllowance(g, pdf) {
  const isTeacher = g.group === "교원" || g.group === "방과후교사";
  const allowanceTerms = isTeacher ? CONFIG.pdfSectionTerms.teacherAllowance : CONFIG.pdfSectionTerms.staffAllowance;
  const itemTerms = [g.item, ...((CONFIG.allowanceColumns.find((a) => a.key === g.key) || {}).aliases || [])];
  const direct = findRelevantBlocks(pdf.lines, itemTerms);
  let pdfAmount = bestAmount(direct.flatMap(extractAmounts), g.amount);
  let moneyStatus = amountStatus(g.amount, pdfAmount);
  let formation = pdfAmount != null ? "단일 편성" : "미편성";
  let note = direct.length ? `개별 수당 항목 확인: ${clip(direct[0])}` : "개별 수당 항목을 찾지 못했습니다.";

  if (pdfAmount == null) {
    const integrated = findLine(pdf.lines, allowanceTerms);
    if (integrated) {
      formation = "통합 편성 의심";
      moneyStatus = "비교불가";
      note = `${g.item}이 개별 항목이 아니라 교원수당/직원수당 등으로 통합 편성되었을 가능성이 있습니다. 확인 위치: ${clip(integrated)}`;
    }
  }
  return finish(g, pdfAmount, moneyStatus, formation, note);
}

function finish(g, pdfAmount, moneyStatus, formation, note) {
  const final = (formation === "미편성" || moneyStatus === "미편성" || moneyStatus === "과소편성" || moneyStatus === "과다편성") ? "부적정" :
    (formation.includes("의심") || formation.includes("급식비") || moneyStatus === "비교불가") ? "확인 필요" : "적정";
  return { group: g.group, item: g.item, excelAmount: g.amount, pdfAmount, diff: pdfAmount == null ? null : pdfAmount - g.amount, moneyStatus, formation, final, note };
}

function baseTerms(group) {
  if (group === "방과후교사") return CONFIG.pdfSectionTerms.afterBase;
  if (group === "영양사") return ["영양사", ...CONFIG.pdfSectionTerms.staffBase, ...CONFIG.pdfSectionTerms.nutritionist];
  if (group === "직원") return CONFIG.pdfSectionTerms.staffBase;
  return CONFIG.pdfSectionTerms.teacherBase;
}

function findRelevantBlocks(lines, terms) {
  const out = [];
  const normTerms = terms.map(normalize);
  for (let i = 0; i < lines.length; i++) {
    const block = lines.slice(Math.max(0, i - 1), Math.min(lines.length, i + 3)).join(" ");
    if (normTerms.some((t) => normalize(block).includes(t))) out.push(block);
  }
  return [...new Set(out)];
}

function findLine(lines, terms) {
  const normTerms = terms.map(normalize);
  return lines.find((l) => normTerms.some((t) => normalize(l).includes(t))) || null;
}

function extractCalcResults(text) {
  return [...String(text).matchAll(/=\s*([\d,]+)\s*원?/g)].map((m) => money(m[1])).filter(Boolean);
}
function extractAmounts(text) {
  return (String(text).match(/[\d,]{4,}\s*(원|천원)?/g) || []).map(money).filter(Boolean).map(normalizePdfAmount);
}
function extractHeadcount(text) {
  const m = String(text).match(/(\d{1,3})\s*(명|인)/);
  return m ? Number(m[1]) : 0;
}
function bestAmount(amounts, expected) {
  if (!amounts.length) return null;
  const scored = amounts.map((a) => normalizePdfAmount(a)).filter(Boolean).map((a) => ({ a, d: Math.abs(Math.floor(a / 1000) - Math.floor(expected / 1000)) })).sort((x, y) => x.d - y.d);
  return scored.length && scored[0].d <= Math.max(3, Math.floor(expected / 1000) * 0.02) ? scored[0].a : null;
}
function amountStatus(excelAmount, pdfAmount) {
  if (pdfAmount == null) return "미편성";
  const e = Math.floor(excelAmount / 1000), p = Math.floor(pdfAmount / 1000);
  if (e === p) return "일치";
  return p < e ? "과소편성" : "과다편성";
}
function sameThousand(a, b) { return Math.floor(normalizePdfAmount(a) / 1000) === Math.floor(normalizePdfAmount(b) / 1000); }
function normalizePdfAmount(n) { n = Number(n || 0); return n > 0 && n < 10000000 ? n * 1000 : n; }
function inferMonth(v) { return v < 10000000 ? 12 : 1; }
function sumVisibleMoney(matrix) { return matrix.flat().map(money).filter((n) => n > 0).reduce((a, b) => a + b, 0); }
function money(v) { const s = String(v ?? "").replace(/\s/g, ""); if (!/\d/.test(s)) return 0; const n = Number(s.replace(/[^0-9.-]/g, "")); return Number.isFinite(n) ? Math.abs(n) : 0; }
function normalize(s) { return String(s || "").replace(/\s+/g, "").toLowerCase(); }
function hasAny(text, patterns) { const n = normalize(text); return patterns.some((p) => n.includes(normalize(p))); }
function clip(s) { return String(s || "").replace(/\s+/g, " ").slice(0, 240); }

function renderResults(rows, debug) {
  const tbody = $("resultsTable").querySelector("tbody");
  tbody.innerHTML = "";
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.group)}</td><td>${escapeHtml(r.item)}</td>
      <td class="amount">${formatWon(r.excelAmount)}</td><td class="amount">${r.pdfAmount == null ? "-" : formatWon(r.pdfAmount)}</td>
      <td class="amount">${r.diff == null ? "-" : formatWon(r.diff)}</td>
      <td>${badge(r.moneyStatus)}</td><td>${escapeHtml(r.formation)}</td><td>${badge(r.final)}</td><td>${escapeHtml(r.note)}</td>`;
    tbody.appendChild(tr);
  }
  renderSummary(rows, debug);
  $("resultsCard").hidden = false; $("summaryCard").hidden = false; $("debugCard").hidden = false;
  $("debugOutput").textContent = JSON.stringify(debug, null, 2);
  $("downloadCsvBtn").disabled = false; $("downloadJsonBtn").disabled = false;
}
function renderSummary(rows, debug) {
  const counts = rows.reduce((acc, r) => { acc[r.final] = (acc[r.final] || 0) + 1; return acc; }, {});
  $("summary").innerHTML = [metric("전체 항목", rows.length), metric("적정", counts["적정"] || 0), metric("확인 필요", counts["확인 필요"] || 0), metric("부적정", counts["부적정"] || 0), metric("제외된 숨김 시트", debug.hiddenSheets.length)].join("");
}
function metric(label, value) { return `<div class="metric"><strong>${value}</strong><span>${label}</span></div>`; }
function badge(text) { const cls = text === "적정" || text === "일치" || text === "편성 여부만 확인" ? "ok" : text === "부적정" || text.includes("과") || text === "미편성" ? "bad" : "warn"; return `<span class="badge ${cls}">${escapeHtml(text)}</span>`; }
function formatWon(n) { return Number(n || 0).toLocaleString("ko-KR") + "원"; }
function escapeHtml(s) { return String(s ?? "").replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c])); }
function setStatus(msg, isError = false) { const el = $("status"); el.textContent = msg; el.style.color = isError ? "#b91c1c" : "#667085"; }
function scoreFinal(final) { return final === "부적정" ? 3 : final === "확인 필요" ? 2 : 1; }
function downloadCsv() {
  const headers = ["그룹", "항목", "엑셀 기준", "PDF 확인", "차이", "금액검증", "편성형태", "최종판정", "비고"];
  const lines = [headers, ...state.rows.map((r) => [r.group, r.item, r.excelAmount, r.pdfAmount ?? "", r.diff ?? "", r.moneyStatus, r.formation, r.final, r.note])].map((row) => row.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","));
  downloadBlob("budget-review-result.csv", "text/csv;charset=utf-8", "\ufeff" + lines.join("\n"));
}
function downloadJson() { downloadBlob("budget-review-result.json", "application/json", JSON.stringify({ rows: state.rows, debug: state.debug }, null, 2)); }
function downloadBlob(filename, type, content) { const blob = new Blob([content], { type }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = filename; a.click(); URL.revokeObjectURL(url); }
