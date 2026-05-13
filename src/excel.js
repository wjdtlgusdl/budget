export async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true, raw: false });
}

export function buildExcelFacts(workbook) {
  const sheets = workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' })
  }));
  const allRows = sheets.flatMap((sheet) => sheet.rows.map((row) => ({ sheet: sheet.name, row })));
  const facts = {
    sheetNames: workbook.SheetNames,
    sheets,
    salaryItems: [],
    retirement: detectRetirement(sheets),
    staffCounts: detectStaffCounts(allRows),
  };

  facts.salaryItems = detectSalaryItems(allRows);
  return facts;
}

function detectSalaryItems(allRows) {
  const candidates = [
    { category: '급여', item: '본봉/기본급', keywords: ['본봉', '기본급'], targetKeywords: ['교원급여', '직원급여'] },
    { category: '급여', item: '방과후교원급여', keywords: ['방과후', '기본급'], targetKeywords: ['방과후교원급여', '방과후교사급여'] },
    { category: '수당', item: '시간외수당', keywords: ['시간외'], targetKeywords: ['시간외수당', '교원수당', '직원수당'] },
    { category: '수당', item: '연구활동비', keywords: ['연구활동'], targetKeywords: ['연구활동비', '교원수당'] },
    { category: '인건비', item: '영양사 인건비', keywords: ['영양사'], targetKeywords: ['영양사', '직원급여', '직원인건비'] },
    { category: '인건비', item: '조리원 인건비', keywords: ['조리원', '조리사'], targetKeywords: ['조리원', '조리사', '직원급여', '직원인건비'] },
    { category: '법정부담금', item: '사회보험부담금', keywords: ['사회보험', '기관부담', '4대보험'], targetKeywords: ['사회보험', '기관부담', '법정부담금'] },
  ];

  const detected = [];
  for (const rule of candidates) {
    const matchedRows = allRows.filter(({ row }) => rowContainsAll(row, rule.keywords));
    const amount = sumLikelyAmounts(matchedRows.map((r) => r.row));
    if (amount > 0 || matchedRows.length > 0) {
      detected.push({ ...rule, amount, sourceRows: matchedRows.slice(0, 5).map((r) => `${r.sheet}: ${r.row.join(' | ')}`) });
    }
  }

  // Generic fallback: rows containing 소계/합계 and wage terms.
  const genericRows = allRows.filter(({ row }) => /소계|합계/.test(row.join(' ')) && /본봉|기본급|급여|수당|인건비/.test(row.join(' ')));
  for (const { sheet, row } of genericRows.slice(0, 20)) {
    const text = row.join(' ');
    const amount = maxNumber(row);
    if (amount > 0 && !detected.some((x) => text.includes(x.item))) {
      detected.push({ category: '기타', item: compactLabel(text), amount, keywords: [compactLabel(text)], targetKeywords: [compactLabel(text)], sourceRows: [`${sheet}: ${text}`] });
    }
  }
  return detected;
}

function detectRetirement(sheets) {
  const retirementSheets = sheets.filter((s) => /퇴직|적립/.test(s.name));
  const rows = retirementSheets.flatMap((s) => s.rows.map((row) => ({ sheet: s.name, row })));
  const amount = sumLikelyAmounts(rows.map((r) => r.row));
  const hasAmount = amount > 0 || rows.some(({ row }) => /적립금액|퇴직적립|퇴직급여/.test(row.join(' ')) && maxNumber(row) > 0);
  return { hasAmount, amount, sourceRows: rows.slice(0, 8).map((r) => `${r.sheet}: ${r.row.join(' | ')}`) };
}

function detectStaffCounts(allRows) {
  const map = {};
  for (const { row } of allRows) {
    const text = row.join(' ');
    const peopleMatch = text.match(/(\d+)\s*명/);
    if (!peopleMatch) continue;
    if (/방과후/.test(text)) map['방과후교사'] = Number(peopleMatch[1]);
    if (/영양사/.test(text)) map['영양사'] = Number(peopleMatch[1]);
    if (/조리/.test(text)) map['조리원'] = Number(peopleMatch[1]);
  }
  return map;
}

function rowContainsAll(row, keywords) {
  const text = row.join(' ');
  return keywords.every((kw) => text.includes(kw));
}

function sumLikelyAmounts(rows) {
  let sum = 0;
  for (const row of rows) {
    const nums = row.map(numberFromCell).filter((n) => n >= 1000);
    if (nums.length) sum += Math.max(...nums);
  }
  return sum;
}

function maxNumber(row) {
  const nums = row.map(numberFromCell).filter((n) => n >= 1000);
  return nums.length ? Math.max(...nums) : 0;
}

function numberFromCell(cell) {
  if (typeof cell === 'number') return cell;
  const text = String(cell || '').replace(/,/g, '');
  const matches = text.match(/-?\d+(?:\.\d+)?/g);
  if (!matches) return 0;
  return Math.max(...matches.map(Number));
}

function compactLabel(text) {
  return String(text).replace(/\s+/g, ' ').slice(0, 40);
}
