export async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true, raw: true, sheetStubs: true });
}

export function buildExcelFacts(workbook) {
  const sheets = workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: true })
  }));
  const allRows = sheets.flatMap((sheet) => sheet.rows.map((row, idx) => ({ sheet: sheet.name, row, rowNumber: idx + 1 })));
  const diagnostics = [];
  const facts = {
    sheetNames: workbook.SheetNames,
    sheets,
    budgetPages: buildBudgetPages(sheets),
    salaryItems: [],
    retirement: detectRetirement(sheets),
    staffCounts: detectStaffCounts(allRows),
    diagnostics,
  };

  const structured = detectFromPayrollSheet(sheets, diagnostics);
  const monthly = detectFromMonthlyPayroll(sheets, diagnostics);
  const heuristic = detectSalaryItems(allRows);
  facts.salaryItems = dedupeItems([...structured, ...monthly, ...heuristic]);

  if (!facts.salaryItems.length) {
    facts.diagnostics.push('검토 항목을 추출하지 못했습니다. 단, 이번 버전은 교직원보수일람표와 월급여 시트를 모두 확인했습니다. 디버그 정보의 시트명/헤더행/소계행을 확인해주세요.');
  } else {
    facts.diagnostics.push(`검토 대상 ${facts.salaryItems.length}건을 추출했습니다.`);
  }
  return facts;
}

function detectFromPayrollSheet(sheets, diagnostics = []) {
  const sheet = sheets.find((s) => /보수일람|교직원보수/.test(s.name));
  if (!sheet) {
    diagnostics.push('교직원보수일람표 시트를 찾지 못했습니다.');
    return [];
  }
  const rows = sheet.rows;
  const headerIdx = findPayrollHeaderIndex(rows);
  if (headerIdx < 0) diagnostics.push(`${sheet.name}: 헤더 행 자동 인식 실패. 고정 열/소계 행 기반 보조 추출을 사용합니다.`);
  const header = buildMergedHeader(rows, headerIdx >= 0 ? headerIdx : 4);
  const col = detectPayrollColumns(header);
  // 이 양식의 실제 위치 보정: E열=본봉, F=시간외, H=관리업무, I=기타, K=연구활동비
  if (col.job < 0) col.job = 1;
  if (col.base < 0) col.base = 4;
  if (col.overtime < 0) col.overtime = 5;
  if (col.management < 0) col.management = 7;
  if (col.other < 0) col.other = 8;
  if (col.research < 0) col.research = 10;

  const teacherSubtotal = findSubtotalRow(rows, ['소계교원']);
  const staffSubtotal = findSubtotalRow(rows, ['소계일반직', '소계직원']);
  const items = [];

  if (teacherSubtotal) items.push(makeItem('급여', '교원급여/본봉(기본급)', amountAt(teacherSubtotal, col.base), ['교원급여', '교원기본급'], [`${sheet.name}: ${rowText(teacherSubtotal)}`]));
  if (staffSubtotal) items.push(makeItem('급여', '직원급여/본봉(기본급)', amountAt(staffSubtotal, col.base), ['직원급여', '직원기본급'], [`${sheet.name}: ${rowText(staffSubtotal)}`]));

  const teacherAllowance = sumCols(teacherSubtotal, [col.overtime, col.management, col.other, col.research].filter((i) => i >= 0));
  if (teacherAllowance > 0) items.push(makeItem('수당', '교원수당(시간외·관리업무·기타·연구활동비)', teacherAllowance, ['교원수당', '시간외수당', '관리업무수당', '기타수당', '연구활동비'], [`${sheet.name}: ${rowText(teacherSubtotal)}`]));
  const staffAllowance = sumCols(staffSubtotal, [col.overtime, col.management, col.other, col.research].filter((i) => i >= 0));
  if (staffAllowance > 0) items.push(makeItem('수당', '직원수당(시간외·관리업무·기타·연구활동비)', staffAllowance, ['직원수당', '시간외수당', '관리업무수당', '기타수당', '연구활동비'], [`${sheet.name}: ${rowText(staffSubtotal)}`]));

  diagnostics.push(`${sheet.name}: 헤더행=${headerIdx >= 0 ? headerIdx + 1 : '보정'}, 본봉열=${colName(col.base)}, 교원소계=${teacherSubtotal ? '있음' : '없음'}, 직원소계=${staffSubtotal ? '있음' : '없음'}, 추출=${items.filter(x => x.amount > 0).length}건`);
  return items.filter((x) => x.amount > 0);
}

function detectFromMonthlyPayroll(sheets, diagnostics = []) {
  const sheet = sheets.find((s) => /월급여/.test(s.name));
  if (!sheet) {
    diagnostics.push('월급여 시트를 찾지 못했습니다.');
    return [];
  }
  const rows = sheet.rows;
  const headerIdx = findPayrollHeaderIndex(rows);
  const header = buildMergedHeader(rows, headerIdx >= 0 ? headerIdx : 5);
  const col = detectPayrollColumns(header);
  if (col.job < 0) col.job = 1;
  if (col.base < 0) col.base = 5;
  if (col.months < 0) col.months = 4;
  if (col.overtime < 0) col.overtime = 6;
  if (col.management < 0) col.management = 9;
  if (col.other < 0) col.other = 10;
  if (col.research < 0) col.research = 12;

  const teacherSubtotal = findSubtotalRow(rows, ['소계교원']);
  const staffSubtotal = findSubtotalRow(rows, ['소계일반직', '소계직원']);
  const dataRows = rows.filter((r) => isPersonRowFlexible(r, col.job));
  const afterschoolRows = dataRows.filter((r) => /방과후/.test(normalizedRow(r)));
  const nutritionRows = dataRows.filter((r) => /영양사/.test(normalizedRow(r)));
  const cookRows = dataRows.filter((r) => /조리사|조리원|조리/.test(normalizedRow(r)));

  const annualFactor = 12;
  const items = [];
  if (teacherSubtotal) items.push(makeItem('급여', '교원급여/본봉(기본급)', amountAt(teacherSubtotal, col.base) * annualFactor, ['교원급여', '교원기본급'], [`${sheet.name}: ${rowText(teacherSubtotal)} × 12월`]));
  if (staffSubtotal) items.push(makeItem('급여', '직원급여/본봉(기본급)', amountAt(staffSubtotal, col.base) * annualFactor, ['직원급여', '직원기본급'], [`${sheet.name}: ${rowText(staffSubtotal)} × 12월`]));
  if (afterschoolRows.length) items.push(makeItem('급여', '방과후교원급여', sumMonthlyAnnual(afterschoolRows, col.base, col.months), ['방과후교원급여', '방과후교사급여', '방과후교사'], sourceRows(sheet.name, afterschoolRows), { expectedCount: afterschoolRows.length }));
  if (nutritionRows.length) items.push(makeItem('인건비', '영양사 인건비', sumMonthlyAnnual(nutritionRows, col.base, col.months), ['영양사', '직원급여', '조리직원급여'], sourceRows(sheet.name, nutritionRows), { expectedCount: nutritionRows.length }));
  if (cookRows.length) items.push(makeItem('인건비', '조리원/조리사 인건비', sumMonthlyAnnual(cookRows, col.base, col.months), ['조리직원급여', '조리사', '조리원'], sourceRows(sheet.name, cookRows), { expectedCount: cookRows.length }));

  const teacherAllowance = sumCols(teacherSubtotal, [col.overtime, col.management, col.other, col.research].filter((i) => i >= 0)) * annualFactor;
  if (teacherAllowance > 0) items.push(makeItem('수당', '교원수당(시간외·관리업무·기타·연구활동비)', teacherAllowance, ['교원수당', '시간외수당', '관리업무수당', '기타수당', '연구활동비'], [`${sheet.name}: ${rowText(teacherSubtotal)} × 12월`]));
  const staffAllowance = sumCols(staffSubtotal, [col.overtime, col.management, col.other, col.research].filter((i) => i >= 0)) * annualFactor;
  if (staffAllowance > 0) items.push(makeItem('수당', '직원수당(시간외·관리업무·기타·연구활동비)', staffAllowance, ['직원수당', '시간외수당', '관리업무수당', '기타수당', '연구활동비'], [`${sheet.name}: ${rowText(staffSubtotal)} × 12월`]));

  diagnostics.push(`${sheet.name}: 헤더행=${headerIdx >= 0 ? headerIdx + 1 : '보정'}, 본봉열=${colName(col.base)}, 교원소계=${teacherSubtotal ? '있음' : '없음'}, 직원소계=${staffSubtotal ? '있음' : '없음'}, 방과후=${afterschoolRows.length}명, 영양사=${nutritionRows.length}명, 조리=${cookRows.length}명, 추출=${items.filter(x => x.amount > 0).length}건`);
  return items.filter((x) => x.amount > 0);
}

function findPayrollHeaderIndex(rows) {
  return rows.findIndex((r, idx) => {
    const text = normalizedRow(r);
    const next = normalizedRow(rows[idx + 1] || []);
    return text.includes('직명') && (text.includes('성명') || text.includes('성 명')) && (text + next).includes('본봉');
  });
}

function buildMergedHeader(rows, idx) {
  const a = rows[idx] || [];
  const b = rows[idx + 1] || [];
  const max = Math.max(a.length, b.length);
  return Array.from({ length: max }, (_, i) => normalizeHeader(`${a[i] || ''}${b[i] || ''}`));
}

function detectPayrollColumns(header) {
  return {
    job: findCol(header, ['직명']),
    months: findCol(header, ['근무', '월수']),
    base: findCol(header, ['본봉']),
    monthlyBase: findCol(header, ['본봉']),
    overtime: findCol(header, ['시간외']),
    duty: findCol(header, ['직급보조']),
    management: findCol(header, ['관리업무']),
    other: findCol(header, ['기타수당']),
    research: findCol(header, ['연구활동']),
    meal: findCol(header, ['식대']),
    holiday: findCol(header, ['명절']),
    teacherDay: findCol(header, ['스승']),
    vacation: findCol(header, ['방학']),
    performance: findCol(header, ['성과']),
  };
}

function buildBudgetPages(sheets) {
  const budget = sheets.find((s) => /세출예산명세서/.test(s.name));
  if (!budget) return [];
  return budget.rows.map((row, idx) => ({ page: `엑셀:${budget.name} ${idx + 1}행`, text: rowText(row) })).filter((p) => p.text.trim());
}

function makeItem(category, item, amount, targetKeywords, sourceRows, extra = {}) {
  return { category, item, amount: Math.round(Number(amount) || 0), keywords: [item], targetKeywords, sourceRows, ...extra };
}

function detectSalaryItems(allRows) {
  const candidates = [
    { category: '법정부담금', item: '교원법정부담금', keywords: ['교원법정부담금'], targetKeywords: ['교원법정부담금', '4대보험', '사학연금'] },
    { category: '법정부담금', item: '직원법정부담금', keywords: ['직원법정부담금'], targetKeywords: ['직원법정부담금', '4대보험', '국민연금'] },
  ];
  const detected = [];
  for (const rule of candidates) {
    const matchedRows = allRows.filter(({ row }) => rowContainsAll(row, rule.keywords));
    const amount = sumLikelyAmounts(matchedRows.map((r) => r.row));
    if (amount > 0 || matchedRows.length > 0) detected.push({ ...rule, amount, sourceRows: matchedRows.slice(0, 5).map((r) => `${r.sheet} ${r.rowNumber}행: ${rowText(r.row)}`) });
  }
  return detected;
}

function detectRetirement(sheets) {
  const retirementSheets = sheets.filter((s) => /퇴직|적립/.test(s.name));
  const rows = retirementSheets.flatMap((s) => s.rows.map((row, idx) => ({ sheet: s.name, row, rowNumber: idx + 1 })));
  const amount = sumLikelyAmounts(rows.map((r) => r.row));
  const hasAmount = amount > 0 || rows.some(({ row }) => /적립금액|퇴직적립|퇴직급여|퇴직금/.test(rowText(row)) && maxNumber(row) > 0);
  return { hasAmount, amount, sourceRows: rows.slice(0, 8).map((r) => `${r.sheet} ${r.rowNumber}행: ${rowText(r.row)}`) };
}

function detectStaffCounts(allRows) {
  const map = {};
  for (const { row } of allRows) {
    const text = rowText(row);
    const peopleMatch = text.match(/(\d+)\s*명/);
    if (!peopleMatch) continue;
    if (/방과후/.test(text)) map['방과후교사'] = Number(peopleMatch[1]);
    if (/영양사/.test(text)) map['영양사'] = Number(peopleMatch[1]);
    if (/조리/.test(text)) map['조리원'] = Number(peopleMatch[1]);
  }
  return map;
}

function findSubtotalRow(rows, labels) {
  return rows.find((r) => {
    const t = normalizedRow(r);
    return labels.some((label) => t.includes(label));
  });
}

function isPersonRowFlexible(row, jobCol) {
  const job = String(row[jobCol] || '').trim();
  if (!job || /직명|소계|합계|번호|작성요령/.test(job)) return false;
  if (/^#/.test(job)) return false;
  return /원장|교사|영양사|조리|사무|실장|기사|차량|환경|돌봄|보조|행정/.test(job);
}

function rowContainsAll(row, keywords) { const text = rowText(row); return keywords.every((kw) => text.includes(kw)); }
function sumMonthlyAnnual(rows, baseCol, monthCol) { return rows.reduce((s, r) => s + amountAt(r, baseCol) * (amountAt(r, monthCol) || 12), 0); }
function sumLikelyAmounts(rows) { let sum = 0; for (const row of rows) { const nums = row.map(numberFromCell).filter((n) => n >= 1000); if (nums.length) sum += Math.max(...nums); } return Math.round(sum); }
function maxNumber(row) { const nums = row.map(numberFromCell).filter((n) => n >= 1000); return nums.length ? Math.max(...nums) : 0; }
function amountAt(row, idx) { return idx >= 0 ? numberFromCell(row[idx]) : 0; }
function sumCols(row, cols) { return row ? cols.reduce((s, c) => s + amountAt(row, c), 0) : 0; }
function sourceRows(sheetName, rows) { return rows.slice(0, 8).map((r) => `${sheetName}: ${rowText(r)}`); }
function numberFromCell(cell) { if (typeof cell === 'number') return Number.isFinite(cell) ? cell : 0; if (cell && typeof cell === 'object' && 'v' in cell) return numberFromCell(cell.v); const text = String(cell || '').replace(/,/g, '').replace(/원/g, ''); const matches = text.match(/-?\d+(?:\.\d+)?/g); if (!matches) return 0; return Math.max(...matches.map(Number)); }
function rowText(row) { return row.map((x) => String(x ?? '').replace(/\s+/g, ' ').trim()).filter(Boolean).join(' | '); }
function normalizedRow(row) { return rowText(row).replace(/\s+/g, '').replace(/[()（）\[\]{}·.,|]/g, ''); }
function normalizeHeader(x) { return String(x || '').replace(/\s+/g, '').replace(/[()（）A-Z]/g, ''); }
function findCol(header, terms) { return header.findIndex((h) => terms.every((t) => h.includes(t))); }
function colName(idx) { let s = ''; idx += 1; while (idx) { const m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); } return s; }

function dedupeItems(items) {
  const seen = new Set();
  const out = [];
  for (const item of items) {
    const key = `${item.category}:${item.item}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(item);
  }
  return out;
}
