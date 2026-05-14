
/* global XLSX, pdfjsLib */

const CONFIG = {
  excludedKeywords: ["법정부담금", "4대보험", "사학연금", "국민연금", "건강보험", "고용보험", "산재보험", "사회보험"],
  salarySheetNameKeywords: ["교직원", "보수", "급여", "봉급", "인건비", "일람"],
  retirementKeywords: ["퇴직"],
  allowedPersonnelSections: ["교원급여", "교원수당", "직원급여", "직원수당", "그밖의인건비", "교원인건비", "직원인건비", "인건비"]
};

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const $ = (id) => document.getElementById(id);
const norm = (s) => String(s ?? "").replace(/\s+/g, "").trim();
const fmt = (n) => Number.isFinite(n) ? Math.round(n).toLocaleString("ko-KR") : "-";
const toNumber = (v) => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const cleaned = String(v ?? "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
};
const hasAny = (text, arr) => arr.some(k => norm(text).includes(norm(k)));
const isExcluded = (text) => hasAny(text, CONFIG.excludedKeywords);

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
  const visibility = Object.fromEntries((wb.Workbook?.Sheets || []).map((s, i) => [wb.SheetNames[i], s.Hidden || 0]));
  const sheets = wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const matrix = ws ? worksheetToMatrix(ws) : [];
    return { name, hidden: visibility[name] || 0, matrix, text: matrix.flat().map(String).join(" ") };
  });
  const salaryCandidates = sheets
    .filter(s => !s.hidden)
    .map(s => ({ sheet: s, score: scoreSalarySheet(s) }))
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score);
  const salarySheet = salaryCandidates[0]?.sheet || null;
  const salary = salarySheet ? parseSalarySheet(salarySheet) : null;

  const retirementSheets = sheets.filter(s => !s.hidden && (hasAny(s.name, CONFIG.retirementKeywords) || hasAny(s.text, CONFIG.retirementKeywords)));
  const retirement = parseRetirement(retirementSheets);
  return { sheets, salarySheet, salary, retirement, salaryCandidates };
}


function worksheetToMatrix(ws) {
  if (!ws || !ws["!ref"]) return [];
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const matrix = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      let v = "";
      if (cell) {
        if (cell.v !== undefined && cell.v !== null) v = cell.v;
        else if (cell.w !== undefined && cell.w !== null) v = cell.w;
        else if (cell.f && cell.v !== undefined) v = cell.v;
      }
      row.push(v);
    }
    matrix.push(row);
  }
  return matrix;
}

function scoreSalarySheet(sheet) {
  let score = 0;
  const name = norm(sheet.name);
  const t = norm(sheet.text);

  // 보이는 시트만 후보가 되며, 같은 양식의 월/년 시트가 함께 있을 때는
  // 반드시 연간 기준 시트를 우선합니다.
  // 예: 교직원보수일람표(년) > 교직원보수일람표 > 교직원보수일람표(월)
  if (name.includes("교직원보수일람표")) score += 80;
  if (name.includes("보수일람표")) score += 60;
  if (name.includes("보수") || name.includes("급여") || name.includes("봉급") || name.includes("인건비") || name.includes("교직원")) score += 20;
  if (name.includes("년") || name.includes("연간") || name.includes("연")) score += 100;
  if (name.includes("월") || name.includes("월급여") || name.includes("매월")) score -= 100;

  ["직명", "본봉", "기본급", "지급액계", "소계교원", "소계일반직", "소계직원"].forEach(k => {
    if (t.includes(norm(k))) score += 10;
  });

  // 숨김 월급여/보조 계산 시트가 우연히 후보가 되지 않도록 감점합니다.
  if (name === "월급여" || name.includes("비월정수당")) score -= 200;
  return score;
}

function parseSalarySheet(sheet) {
  const m = sheet.matrix || [];
  const maxCols = Math.max(...m.map(r => r.length), 0);
  const scanRows = Math.min(40, m.length);
  const cell = (r, c) => String((m[r] || [])[c] ?? "");
  const cellNorm = (r, c) => norm(cell(r, c));
  const rowNorm = (r) => norm((m[r] || []).join(" "));
  const colNorm = (c, r1 = 0, r2 = scanRows) => {
    const parts = [];
    for (let r = r1; r < Math.min(r2, m.length); r++) parts.push(cell(r, c));
    return norm(parts.join(" "));
  };

  let jobCol = -1, baseCol = -1, totalCol = -1, headerIndex = -1;

  // 1) 상단 셀 단위 탐색. 줄바꿈/괄호/영문기호는 norm()이 제거합니다.
  for (let r = 0; r < scanRows; r++) {
    for (let c = 0; c < maxCols; c++) {
      const t = cellNorm(r, c);
      if (jobCol < 0 && t.includes("직명")) { jobCol = c; headerIndex = Math.max(headerIndex, r); }
      if (baseCol < 0 && (t.includes("본봉") || t.includes("기본급"))) { baseCol = c; headerIndex = Math.max(headerIndex, r); }
      if (totalCol < 0 && t.includes("지급액계")) { totalCol = c; headerIndex = Math.max(headerIndex, r); }
    }
  }

  // 2) 열 단위 탐색. 병합셀 또는 2단 헤더 보정.
  for (let c = 0; c < maxCols; c++) {
    const t = colNorm(c);
    if (jobCol < 0 && t.includes("직명")) jobCol = c;
    if (baseCol < 0 && (t.includes("본봉") || t.includes("기본급"))) baseCol = c;
    if (totalCol < 0 && t.includes("지급액계")) totalCol = c;
  }

  // 3) 교직원보수일람표 표준 양식 보정: B=직명, E=본봉, Q=지급액계.
  //    실제 예시처럼 XLSX 라이브러리가 헤더 텍스트를 일부 놓치는 경우에도 중단하지 않습니다.
  const titleText = norm(m.slice(0, scanRows).flat().join(" ") + " " + sheet.name);
  if (titleText.includes("교직원보수일람표") || titleText.includes("보수일람표") || titleText.includes("교직원보수")) {
    if (jobCol < 0) jobCol = 1;
    if (baseCol < 0) baseCol = 4;
    if (totalCol < 0) totalCol = 16;
  }

  // 4) 여전히 못 찾은 경우: 상단 20행에서 숫자 데이터 패턴으로 보정.
  //    직명열은 문자열, 본봉열은 큰 금액, 지급액계는 본봉보다 큰 금액이 반복되는 열입니다.
  if (jobCol < 0 || baseCol < 0 || totalCol < 0) {
    for (let c = 0; c < maxCols; c++) {
      const sampleVals = [];
      for (let r = 0; r < Math.min(m.length, 80); r++) sampleVals.push(cell(r, c));
      const txtCount = sampleVals.filter(v => /원장|교사|직원|기사|조리|방과후|미화|보조/.test(String(v))).length;
      const moneyCount = sampleVals.filter(v => toNumber(v) >= 1000000).length;
      if (jobCol < 0 && txtCount >= 3) jobCol = c;
      if (baseCol < 0 && moneyCount >= 5) baseCol = c;
    }
    if (totalCol < 0 && baseCol >= 0) totalCol = Math.min(maxCols - 1, baseCol + 12);
  }

  if (jobCol < 0 || baseCol < 0 || totalCol < 0) {
    const preview = m.slice(0, 12).map((row, i) => `${i + 1}: ${row.map(v => String(v || "").replace(/\n/g, "/")).join(" | ")}`).join("\n");
    throw new Error(`보수 시트 후보(${sheet.name})에서 직명/본봉/지급액계 헤더를 찾지 못했습니다.\n상단 미리보기:\n${preview}`);
  }

  // 헤더 행은 직명 셀이 있는 행, 없으면 본봉/지급액계가 가장 가까운 행, 그래도 없으면 표준 양식 5행(0-index 4)
  for (let r = 0; r < scanRows; r++) {
    if (cellNorm(r, jobCol).includes("직명") || cellNorm(r, baseCol).includes("본봉") || cellNorm(r, totalCol).includes("지급액계")) {
      headerIndex = r; break;
    }
  }
  if (headerIndex < 0) headerIndex = 4;

  const headers = Array.from({ length: maxCols }, (_, c) => {
    const parts = [];
    for (let r = Math.max(0, headerIndex - 1); r <= Math.min(m.length - 1, headerIndex + 2); r++) parts.push(cell(r, c));
    return norm(parts.join(" "));
  });

  const allowanceCols = [];
  for (let c = baseCol + 1; c < totalCol; c++) {
    const rawHeader = headers[c] || colNorm(c, Math.max(0, headerIndex - 1), Math.min(m.length, headerIndex + 3));
    const h = prettyHeader(rawHeader || `열${c + 1}`);
    if (!h || isExcluded(h)) continue;
    allowanceCols.push({ index: c, name: h });
  }

  const rows = [];
  let section = "teacher";
  let dataStarted = false;
  for (let r = headerIndex + 1; r < m.length; r++) {
    const row = m[r] || [];
    const nt = rowNorm(r);
    if (!nt) continue;
    if (nt.includes("소계교원") || nt.includes("소계(교원")) { section = "staff"; dataStarted = true; continue; }
    if (nt.includes("소계직원") || nt.includes("소계일반직") || nt.includes("소계(직원") || nt.includes("소계(일반직") || nt.includes("합계교원일반직") || nt.includes("작성요령")) break;

    const job = String(row[jobCol] || "").trim();
    const name = String(row[jobCol + 1] || "").trim();
    const base = toNumber(row[baseCol]);
    const allowanceValues = allowanceCols.map(c => ({ name: c.name, amount: toNumber(row[c.index]) })).filter(x => x.amount > 0);
    const hasMoney = base > 0 || allowanceValues.length > 0;
    if (!job || !hasMoney) continue;
    if (norm(job).includes("직명") || nt.includes("일련번호") || nt.includes("소계")) continue;
    dataStarted = true;
    rows.push({ rowNumber: r + 1, section, job, name, base, allowances: allowanceValues, total: toNumber(row[totalCol]) });
  }

  if (!rows.length) {
    const preview = m.slice(Math.max(0, headerIndex - 2), headerIndex + 10).map((row, i) => `${Math.max(0, headerIndex - 2) + i + 1}: ${row.map(v => String(v || "").replace(/\n/g, "/")).join(" | ")}`).join("\n");
    throw new Error(`보수 시트(${sheet.name})에서 데이터 행을 찾지 못했습니다. 인식 열: 직명 ${jobCol + 1}, 본봉 ${baseCol + 1}, 지급액계 ${totalCol + 1}\n미리보기:\n${preview}`);
  }

  const summary = buildSalarySummary(rows, allowanceCols);
  summary.detectedColumns = { jobCol: jobCol + 1, baseCol: baseCol + 1, totalCol: totalCol + 1, headerRow: headerIndex + 1 };
  return summary;
}

function prettyHeader(h) {
  return String(h || "")
    .replace(/[\n\r]/g, "")
    .replace(/\([A-Z]\)/g, "")
    .replace(/[A-Z]\)?/g, "")
    .replace(/[()=~]/g, " ")
    .replace(/\s+/g, "")
    .trim();
}

function buildSalarySummary(rows, allowanceCols) {
  const teachers = rows.filter(r => r.section === "teacher");
  const regularTeachers = teachers.filter(r => !norm(r.job).includes("방과후"));
  const afterSchoolTeachers = teachers.filter(r => norm(r.job).includes("방과후"));
  const staff = rows.filter(r => r.section === "staff");
  const sumBase = arr => arr.reduce((a, r) => a + r.base, 0);
  const collect = arr => {
    const map = new Map();
    arr.forEach(r => r.allowances.forEach(x => map.set(x.name, (map.get(x.name) || 0) + x.amount)));
    return [...map.entries()].map(([name, amount]) => ({ name, amount })).filter(x => x.amount > 0);
  };
  return {
    rows,
    teachers,
    regularTeachers,
    afterSchoolTeachers,
    staff,
    regularTeacherPay: { amount: sumBase(regularTeachers), people: regularTeachers.filter(r => r.base > 0).length },
    afterSchoolPay: { amount: sumBase(afterSchoolTeachers), people: afterSchoolTeachers.filter(r => r.base > 0).length },
    teacherPayTotal: { amount: sumBase(teachers), people: teachers.filter(r => r.base > 0).length },
    staffPay: { amount: sumBase(staff), people: staff.filter(r => r.base > 0).length },
    teacherAllowances: collect(teachers),
    regularTeacherAllowances: collect(regularTeachers),
    staffAllowances: collect(staff),
    allowanceHeaders: allowanceCols.map(c => c.name)
  };
}

function parseRetirement(sheets) {
  let amount = 0;
  const used = [];
  sheets.forEach(s => {
    let local = 0;
    for (const row of s.matrix) {
      for (const v of row) {
        // 퇴직 관련 시트라도 수식 결과 또는 입력값이 0이면
        // 퇴직 적립금액이 없는 것으로 봅니다.
        // XLSX가 수식 셀을 객체로 넘기는 경우에는 계산 결과(v.v)를 우선 사용합니다.
        let raw = v;
        if (raw && typeof raw === "object" && "v" in raw) raw = raw.v;
        const n = toNumber(raw);
        if (n > 0) local += n;
      }
    }
    if (local > 0) { amount += local; used.push(s.name); }
  });
  return { hasAmount: amount > 0, amount, sheets: used };
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
      .map(([, arr]) => arr.sort((a, b) => a.x - b.x).map(x => x.str).join(" ").replace(/\s+/g, " ").trim())
      .filter(Boolean);
    pages.push({ page: i, lines, text: lines.join("\n") });
  }
  const allLines = pages.flatMap(p => p.lines.map(line => ({ page: p.page, line })));
  return { pages, lines: allLines, fullText: pages.map(p => p.text).join("\n"), compactText: norm(pages.map(p => p.text).join("\n")), ...parsePdfBudget(allLines) };
}

function parsePdfBudget(lines) {
  const totals = {};
  const details = [];
  let currentSection = "";
  const totalRowRe = /^\s*([가-힣A-Za-z0-9(),·]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)/;
  const formulaRe = /([0-9,]+)원\s*\*\s*(?:(\d+)명\s*\*\s*)?([0-9]+)월\s*=\s*([0-9,]+)/;
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i].line;
    const line = raw.replace(/\s+/g, " ").trim();
    const tm = line.match(totalRowRe);
    if (tm && !line.includes("예산구분")) {
      const item = cleanPdfName(tm[1]);
      if (item && !["관", "항", "목"].includes(item)) {
        const amount = toNumber(tm[5]) * 1000;
        totals[item] = { item, amount, page: lines[i].page, raw: line };
        currentSection = item;
      }
    }
    const fm = line.match(formulaRe);
    if (fm) {
      const item = inferDetailName(lines, i);
      details.push({
        item,
        section: currentSection,
        unit: toNumber(fm[1]),
        people: fm[2] ? Number(fm[2]) : 0,
        months: Number(fm[3]),
        amount: toNumber(fm[4]),
        page: lines[i].page,
        raw: line
      });
    }
  }
  return { totals, details };
}

function inferDetailName(lines, i) {
  const currentPage = lines[i].page;

  const candidates = [];
  // 같은 줄에서 산출식 앞쪽에 항목명이 붙어 있는 경우
  const beforeFormula = lines[i].line.split(/\d[\d,]*원\s*\*/)[0] || "";
  if (beforeFormula) candidates.push(beforeFormula);

  // 보통 PDF 텍스트는 “(본예산) ... 항목명” 다음 줄에 산출식이 옵니다.
  for (let j = i - 1; j >= Math.max(0, i - 8); j--) {
    if (lines[j].page !== currentPage) break;
    candidates.push(lines[j].line);
  }
  for (let j = i + 1; j < Math.min(lines.length, i + 4); j++) {
    if (lines[j].page !== currentPage) break;
    candidates.push(lines[j].line);
  }

  for (const cand of candidates) {
    let s = String(cand || "").trim();
    if (!s || /원\s*\*/.test(s) || s.includes("발행일") || s.includes("예산구분")) continue;
    // 총액 행은 산출항목명이 아니므로 제외
    if (/^[가-힣A-Za-z0-9(),·]+\s+[0-9,]+\s+[0-9,]+\s+[0-9,]+\s+[0-9,]+/.test(s)) continue;
    let name = cleanPdfName(s);
    name = repairPdfItemName(name);
    if (name && name.length >= 3 && !["본예산", "보조금및지원금", "수익자부담금"].includes(name)) return name;
  }
  return "산출항목명확인불가";
}

function repairPdfItemName(name) {
  return cleanPdfName(name)
    .replace(/^원급여$/, "교원급여")
    .replace(/^교원급여교원급여$/, "교원급여")
    .replace(/^원수당$/, "교원수당")
    .replace(/^경미화원급여$/, "환경미화원급여")
    .replace(/^량기사급여$/, "차량기사급여")
    .replace(/^량적립금$/, "차량적립금")
    .replace(/^리대체인력비$/, "조리대체인력비")
    .replace(/^품구입비$/, "비품구입비")
    .replace(/^수공사비$/, "보수공사비");
}

function cleanPdfName(s) {
  return norm(String(s || "").replace(/[()]/g, "").replace(/본예산|보조금및지원금|수익자부담금|그밖의수입/g, "").replace(/[^가-힣A-Za-z0-9]/g, ""));
}

function analyze(excel, pdf) {
  const results = [];
  const issues = [];
  const notes = [];
  const debug = [];
  if (!excel.salary) throw new Error("보수/급여 관련 시트를 찾지 못했거나 표 구조를 인식하지 못했습니다.");
  const s = excel.salary;

  debug.push(`선택된 보수 시트: ${excel.salarySheet?.name || "없음"}`);
  debug.push(`보수 인식: 교원 ${s.teachers.length}명(정규 ${s.regularTeachers.length}, 방과후 ${s.afterSchoolTeachers.length}), 직원 ${s.staff.length}명`);
  debug.push(`인식 열: 직명 ${s.detectedColumns?.jobCol || "?"}, 본봉 ${s.detectedColumns?.baseCol || "?"}, 지급액계 ${s.detectedColumns?.totalCol || "?"}, 헤더행 ${s.detectedColumns?.headerRow || "?"}`);
  debug.push(`수당 열: ${s.allowanceHeaders.join(", ")}`);
  debug.push(`퇴직 관련 시트: ${excel.retirement.sheets.join(", ") || "없음"}`);

  compareGroup(results, "보수", "교원급여", { amount: s.teacherPayTotal.amount, people: s.regularTeacherPay.people }, getPdfTeacherPayGroup(pdf), true);
  compareGroup(results, "보수", "방과후교원급여", s.afterSchoolPay, getPdfDetailSum(pdf, ["방과후교원급여"]), true);
  compareGroup(results, "보수", "직원급여", s.staffPay, getPdfTotal(pdf, "직원급여"), false);
  compareAllowance(results, notes, "교원수당", s.teacherAllowances, pdf);
  compareAllowance(results, notes, "직원수당", s.staffAllowances, pdf);

  checkSelfDrivingAllowance(issues, s, pdf);
  checkMisclassifiedPersonnel(issues, pdf);
  checkRetirement(issues, excel, pdf);

  const all = [...results, ...issues];
  return { results: all, debug: debug.concat(buildPdfDebug(pdf), notes) };
}

function getPdfTotal(pdf, item) {
  const found = Object.values(pdf.totals).find(x => norm(x.item) === norm(item));
  return found ? { amount: found.amount, people: 0, basis: found.raw, page: found.page } : { amount: 0, people: 0, basis: "", page: 0 };
}

function getPdfDetailSum(pdf, names) {
  const arr = pdf.details.filter(d => names.some(n => norm(d.item).includes(norm(n))));
  return { amount: arr.reduce((a, d) => a + d.amount, 0), people: uniquePeople(arr), basis: arr.map(d => `${d.item} ${fmt(d.amount)}원`).join(" + "), page: arr[0]?.page || 0 };
}

function getPdfTeacherPayGroup(pdf) {
  const total = getPdfTotal(pdf, "교원급여");
  const arr = pdf.details.filter(d => norm(d.item).includes("교원급여") && !norm(d.item).includes("방과후"));
  return { amount: total.amount || arr.reduce((a, d) => a + d.amount, 0), people: uniquePeople(arr), basis: arr.map(d => `${d.item} ${fmt(d.amount)}원/${d.people || "?"}명`).join(" + "), page: total.page || arr[0]?.page || 0 };
}

function uniquePeople(arr) {
  const vals = arr.map(d => d.people).filter(n => n > 0);
  if (!vals.length) return 0;
  const uniq = [...new Set(vals)];
  return uniq.length === 1 ? uniq[0] : vals.reduce((a, b) => a + b, 0);
}

function compareGroup(results, category, item, excel, pdf, checkPeople) {
  const amountDiff = (excel.amount || 0) - (pdf.amount || 0);
  const peopleDiff = (excel.people || 0) - (pdf.people || 0);
  let status = "ok";
  const msgs = [];
  if (Math.abs(amountDiff) <= 1000) msgs.push("금액 일치");
  else { status = "bad"; msgs.push(`금액 차이 ${fmt(amountDiff)}원`); }
  if (checkPeople && pdf.people > 0) {
    if (peopleDiff === 0) msgs.push("인원 일치");
    else { status = "bad"; msgs.push(`인원수 불일치: PDF ${pdf.people}명, 엑셀 ${excel.people}명`); }
  }
  results.push({ category, item, excelAmount: excel.amount, pdfAmount: pdf.amount, excelPeople: excel.people, pdfPeople: pdf.people || "-", status, message: msgs.join(" / ") + (pdf.page ? ` (PDF ${pdf.page}쪽)` : "") });
}

function compareAllowance(results, notes, bucket, allowances, pdf) {
  const pdfTotal = getPdfTotal(pdf, bucket);
  const direct = [];
  const missing = [];
  allowances.forEach(a => {
    if (isExcluded(a.name)) return;
    const found = pdf.details.find(d => norm(d.item).includes(norm(a.name)) || norm(a.name).includes(norm(d.item)));
    if (found && Math.abs(found.amount - a.amount) <= 1000) direct.push({ ...a, pdfAmount: found.amount });
    else missing.push(a);
  });
  const excelTotal = allowances.reduce((a, x) => a + x.amount, 0);
  const diff = excelTotal - pdfTotal.amount;
  let status = Math.abs(diff) <= 1000 ? "ok" : "warn";
  let message = Math.abs(diff) <= 1000 ? "수당 총액 일치" : `수당 총액 차이 ${fmt(diff)}원`;
  const missingSum = missing.reduce((a, x) => a + x.amount, 0);
  const bucketDetail = pdf.details.find(d => norm(d.item) === norm(bucket));
  if (bucketDetail && Math.abs(missingSum - bucketDetail.amount) <= 1000) {
    message += ` / ${bucket}에 통합편성 추정 (${missing.map(x => `${x.name} ${fmt(x.amount)}원`).join(" + ")} = ${fmt(missingSum)}원)`;
  } else if (missing.length) {
    message += ` / 개별 미확인 수당 ${missing.length}건은 추가확인필요`;
  }
  results.push({ category: "보수", item: bucket, excelAmount: excelTotal, pdfAmount: pdfTotal.amount, excelPeople: "-", pdfPeople: "-", status, message });
  if (direct.length) notes.push(`${bucket} 직접 확인: ${direct.map(x => `${x.name} ${fmt(x.amount)}원`).join(", ")}`);
}

function checkSelfDrivingAllowance(issues, salary, pdf) {
  const teacher = salary.teacherAllowances.filter(a => norm(a.name).includes("자가운전"));
  const staff = salary.staffAllowances.filter(a => norm(a.name).includes("자가운전"));
  const pdfHas = pdf.details.some(d => norm(d.item).includes("자가운전"));
  const total = [...teacher, ...staff].reduce((a, x) => a + x.amount, 0);
  if (total > 0 && !pdfHas) {
    issues.push({ category: "지적사항", item: "자가운전보조금", excelAmount: total, pdfAmount: 0, excelPeople: "-", pdfPeople: "-", status: "bad", message: `교원 및 직원 자가운전보조금 예산 미편성 (엑셀 ${fmt(total)}원)` });
  }
}

function checkMisclassifiedPersonnel(issues, pdf) {
  // 1차 버전에서는 오탐을 막기 위해 구체적으로 확인된 사례만 지적합니다.
  // 향후 영양사/조리원 등은 “항목명+금액+비인건비 위치”가 안정적으로 잡힐 때 확장합니다.
  pdf.details.forEach(d => {
    const item = norm(d.item);
    const section = norm(d.section);
    if (item.includes("차량기사급여") && section.includes("통학차량이용비")) {
      issues.push({ category: "지적사항", item: "차량기사급여 오편성", excelAmount: "-", pdfAmount: d.amount, excelPeople: "-", pdfPeople: d.people || "-", status: "bad", message: `차량기사급여 ${fmt(d.amount)}원을 직원인건비가 아닌 통학차량이용비의 산출내역에 편성 (PDF ${d.page}쪽)` });
    }
  });
}

function checkRetirement(issues, excel, pdf) {
  const excelHas = excel.retirement.hasAmount;
  const pdfHas = /퇴직|퇴직적립|퇴직금|퇴직급여충당/.test(pdf.fullText);
  if (excelHas && !pdfHas) {
    issues.push({ category: "지적사항", item: "퇴직적립금", excelAmount: "있음", pdfAmount: "없음", excelPeople: "-", pdfPeople: "-", status: "bad", message: `엑셀에는 퇴직 적립금액이 있으나 퇴직적립금 미편성 (퇴직 관련 시트: ${excel.retirement.sheets.join(", ")})` });
  } else if (excelHas && pdfHas) {
    issues.push({ category: "확인", item: "퇴직적립금", excelAmount: "있음", pdfAmount: "있음", excelPeople: "-", pdfPeople: "-", status: "ok", message: "엑셀 퇴직 적립금액 및 PDF 퇴직 관련 편성 확인" });
  }
}

function buildPdfDebug(pdf) {
  const lines = [];
  lines.push(`PDF 총액 행: ${Object.keys(pdf.totals).join(", ")}`);
  lines.push("PDF 산출기초:");
  pdf.details.forEach(d => lines.push(`- [${d.section || "구간미상"}] ${d.item}: ${fmt(d.amount)}원, ${d.people || "-"}명, PDF ${d.page}쪽`));
  return lines;
}

function renderReport(report) {
  const results = report.results;
  const summary = $("summary");
  summary.hidden = false;
  const bad = results.filter(r => r.status === "bad").length;
  const warn = results.filter(r => r.status === "warn").length;
  const ok = results.filter(r => r.status === "ok").length;
  summary.innerHTML = `<div class="summary-card bad"><strong>${bad}</strong><span>지적</span></div><div class="summary-card warn"><strong>${warn}</strong><span>확인</span></div><div class="summary-card ok"><strong>${ok}</strong><span>정상</span></div>`;

  const el = $("results");
  if (!results.length) { el.textContent = "표시할 결과가 없습니다."; return; }
  el.innerHTML = `<table><thead><tr><th>구분</th><th>항목</th><th>엑셀 금액</th><th>PDF 금액</th><th>엑셀 인원</th><th>PDF 인원</th><th>판정</th><th>내용</th></tr></thead><tbody>${results.map(r => `<tr><td>${esc(r.category)}</td><td>${esc(r.item)}</td><td>${typeof r.excelAmount === "number" ? fmt(r.excelAmount) + "원" : esc(r.excelAmount)}</td><td>${typeof r.pdfAmount === "number" ? fmt(r.pdfAmount) + "원" : esc(r.pdfAmount)}</td><td>${esc(r.excelPeople)}</td><td>${esc(r.pdfPeople)}</td><td><span class="pill ${r.status}">${r.status === "bad" ? "지적" : r.status === "warn" ? "확인" : "정상"}</span></td><td>${esc(r.message)}</td></tr>`).join("")}</tbody></table>`;
  $("debugLog").textContent = report.debug.join("\n");
}

function esc(v) {
  return String(v ?? "-").replace(/[&<>"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
}
