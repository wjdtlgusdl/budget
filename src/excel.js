export async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true, raw: false });
}

export function buildExcelFacts(workbook) {
  const sheets = workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' })
  }));
  const allRows = sheets.flatMap((sheet) => sheet.rows.map((row, idx) => ({ sheet: sheet.name, row, rowNumber: idx + 1 })));
  const facts = {
    sheetNames: workbook.SheetNames,
    sheets,
    budgetPages: buildBudgetPages(sheets),
    salaryItems: [],
    retirement: detectRetirement(sheets),
    staffCounts: detectStaffCounts(allRows),
    diagnostics: [],
  };

  const structured = detectFromPayrollSheet(sheets);
  const heuristic = detectSalaryItems(allRows);
  facts.salaryItems = dedupeItems([...structured, ...heuristic]);

  if (!facts.salaryItems.length) {
    facts.diagnostics.push('교직원보수일람표에서 검토 대상 급여·수당 항목을 찾지 못했습니다. 시트명 또는 양식 구조를 확인해주세요.');
  }
  return facts;
}

function detectFromPayrollSheet(sheets) {
  const sheet = sheets.find((s) => /보수일람|교직원보수/.test(s.name));
  if (!sheet) return [];
  const rows = sheet.rows;
  const headerIdx = rows.findIndex((r) => /직명/.test(rowText(r)) && /본봉|기본급/.test(rowText(r)));
  if (headerIdx < 0) return [];
  const header = rows[headerIdx].map((x) => normalizeHeader(x));
  const col = {
    job: findCol(header, ['직명']),
    base: findCol(header, ['본봉', '기본급']),
    overtime: findCol(header, ['시간외수당']),
    duty: findCol(header, ['직급보조비']),
    management: findCol(header, ['관리업무', '수당']),
    other: findCol(header, ['기타수당']),
    research: findCol(header, ['연구활동비']),
    meal: findCol(header, ['식대']),
    holiday: findCol(header, ['명절']),
    teacherDay: findCol(header, ['스승']),
    vacation: findCol(header, ['방학']),
    performance: findCol(header, ['성과']),
  };

  const teacherSubtotal = rows.find((r) => /소계\s*\(?교원\)?/.test(rowText(r)));
  const staffSubtotal = rows.find((r) => /소계\s*\(?일반직\)?|소계\s*\(?직원\)?/.test(rowText(r)));
  const dataRows = rows.slice(headerIdx + 1).filter((r) => !/소계|합계|작성|수당란/.test(String(r[0] || '')));
  const afterschoolRows = dataRows.filter((r) => /방과후교사/.test(rowText(r)));
  const nutritionRows = dataRows.filter((r) => /영양사/.test(rowText(r)));
  const cookRows = dataRows.filter((r) => /조리사|조리원/.test(rowText(r)));

  const items = [];
  if (teacherSubtotal && col.base >= 0) items.push(makeItem('급여', '교원급여/본봉(기본급)', amountAt(teacherSubtotal, col.base), ['교원급여', '교원기본급'], [`${sheet.name}: ${rowText(teacherSubtotal)}`]));
  if (staffSubtotal && col.base >= 0) items.push(makeItem('급여', '직원급여/본봉(기본급)', amountAt(staffSubtotal, col.base), ['직원급여', '직원기본급'], [`${sheet.name}: ${rowText(staffSubtotal)}`]));
  if (afterschoolRows.length && col.base >= 0) items.push(makeItem('급여', '방과후교원급여', sumCol(afterschoolRows, col.base), ['방과후교원급여', '방과후교사급여', '방과후교사'], sourceRows(sheet.name, afterschoolRows), { expectedCount: afterschoolRows.length }));
  if (nutritionRows.length && col.base >= 0) items.push(makeItem('인건비', '영양사 인건비', sumCol(nutritionRows, col.base), ['영양사', '직원급여', '조리직원급여'], sourceRows(sheet.name, nutritionRows), { expectedCount: nutritionRows.length }));
  if (cookRows.length && col.base >= 0) items.push(makeItem('인건비', '조리원/조리사 인건비', sumCol(cookRows, col.base), ['조리직원급여', '조리사', '조리원'], sourceRows(sheet.name, cookRows), { expectedCount: cookRows.length }));

  const teacherAllowance = sumCols(teacherSubtotal, [col.overtime, col.management, col.other, col.research].filter((i) => i >= 0));
  if (teacherAllowance > 0) items.push(makeItem('수당', '교원수당(시간외·관리업무·기타·연구활동비)', teacherAllowance, ['교원수당', '시간외수당', '관리업무수당', '기타수당', '연구활동비'], [`${sheet.name}: ${rowText(teacherSubtotal)}`]));
  const staffAllowance = sumCols(staffSubtotal, [col.overtime, col.management, col.other, col.research].filter((i) => i >= 0));
  if (staffAllowance > 0) items.push(makeItem('수당', '직원수당(시간외·관리업무·기타·연구활동비)', staffAllowance, ['직원수당', '시간외수당', '관리업무수당', '기타수당', '연구활동비'], [`${sheet.name}: ${rowText(staffSubtotal)}`]));
  return items.filter((x) => x.amount > 0);
}

function buildBudgetPages(sheets) {
  const budget = sheets.find((s) => /세출예산명세서/.test(s.name));
  if (!budget) return [];
  return budget.rows.map((row, idx) => ({ page: `엑셀:${budget.name} ${idx + 1}행`, text: rowText(row) })).filter((p) => p.text.trim());
}

function makeItem(category, item, amount, targetKeywords, sourceRows, extra = {}) {
  return { category, item, amount: Math.round(amount), keywords: [item], targetKeywords, sourceRows, ...extra };
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
    if (amount > 0 || matchedRows.length > 0) {
      detected.push({ ...rule, amount, sourceRows: matchedRows.slice(0, 5).map((r) => `${r.sheet} ${r.rowNumber}행: ${rowText(r.row)}`) });
    }
  }
  return detected;
}

function detectRetirement(sheets) {
  const retirementSheets = sheets.filter((s) => /퇴직|적립/.test(s.name));
  const rows = retirementSheets.flatMap((s) => s.rows.map((row, idx) => ({ sheet: s.name, row, rowNumber: idx + 1 })));
  const amount = sumLikelyAmounts(rows.map((r) => r.row));
  const hasAmount = amount > 0 || rows.some(({ row }) => /적립금액|퇴직적립|퇴직급여/.test(rowText(row)) && maxNumber(row) > 0);
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

function rowContainsAll(row, keywords) {
  const text = rowText(row);
  return keywords.every((kw) => text.includes(kw));
}

function sumLikelyAmounts(rows) {
  let sum = 0;
  for (const row of rows) {
    const nums = row.map(numberFromCell).filter((n) => n >= 1000);
    if (nums.length) sum += Math.max(...nums);
  }
  return Math.round(sum);
}

function maxNumber(row) {
  const nums = row.map(numberFromCell).filter((n) => n >= 1000);
  return nums.length ? Math.max(...nums) : 0;
}

function amountAt(row, idx) { return idx >= 0 ? numberFromCell(row[idx]) : 0; }
function sumCol(rows, idx) { return rows.reduce((s, r) => s + amountAt(r, idx), 0); }
function sumCols(row, cols) { return row ? cols.reduce((s, c) => s + amountAt(row, c), 0) : 0; }
function sourceRows(sheetName, rows) { return rows.slice(0, 8).map((r) => `${sheetName}: ${rowText(r)}`); }

function numberFromCell(cell) {
  if (typeof cell === 'number') return cell;
  const text = String(cell || '').replace(/,/g, '');
  const matches = text.match(/-?\d+(?:\.\d+)?/g);
  if (!matches) return 0;
  return Math.max(...matches.map(Number));
}

function rowText(row) { return row.map((x) => String(x ?? '').replace(/\s+/g, ' ').trim()).filter(Boolean).join(' | '); }
function normalizeHeader(x) { return String(x || '').replace(/\s+/g, '').replace(/[()（）A-Z]/g, ''); }
function findCol(header, terms) { return header.findIndex((h) => terms.every((t) => h.includes(t))); }

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
