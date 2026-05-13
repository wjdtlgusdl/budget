import { normalizeText } from './pdf.js';

const HUMAN_COST_HINTS = ['급여', '수당', '인건비', '보수', '퇴직', '사회보험', '법정부담'];
const WRONG_BUCKET_HINTS = ['급식비', '급식운영', '운영비', '교육활동비', '관리운영비', '사업비'];

export function runReview(excelFacts, pdfText, options) {
  const results = [];
  const reviewSource = mergeReviewSources(pdfText, excelFacts);
  for (const item of excelFacts.salaryItems) {
    results.push(reviewSalaryItem(item, reviewSource, options));
  }
  if (!excelFacts.salaryItems.length) {
    results.push({
      category: '진단', item: '검토 항목 추출', excelAmount: null, pdfAmount: null, diff: null, status: '확인필요',
      reason: (excelFacts.diagnostics || []).join(' ') || '검토 대상 항목이 추출되지 않았습니다.', evidence: ''
    });
  }
  if (excelFacts.retirement.hasAmount) {
    results.push(reviewRetirement(excelFacts.retirement, pdfText, options));
  }
  return prioritize(results);
}

function mergeReviewSources(pdfText, excelFacts) {
  const pages = [
    ...(pdfText?.pages || []),
    ...(excelFacts?.budgetPages || []),
  ];
  return { raw: `${pdfText?.raw || ''}\n${pages.map((p) => p.text).join('\n')}`, pages };
}

function reviewSalaryItem(item, pdfText, options) {
  const normalMatches = findMatches(pdfText, item.targetKeywords, HUMAN_COST_HINTS);
  const genericMatches = findMatches(pdfText, item.targetKeywords, []);
  const best = normalMatches[0] || genericMatches[0];

  if (!best) {
    return {
      category: item.category,
      item: item.item,
      excelAmount: item.amount || null,
      pdfAmount: null,
      diff: null,
      status: '미편성',
      reason: '세출예산명세서에서 관련 편성 항목을 확인하지 못했습니다.',
      evidence: sourceEvidence(item),
    };
  }

  const pdfAmount = inferPdfAmount(best.text, item.amount, options.allowTruncation);
  const wrongBucket = options.searchWholePdf && WRONG_BUCKET_HINTS.some((kw) => best.text.includes(kw)) && !/급여|수당|인건비|법정부담|퇴직/.test(best.text);
  const integratedAllowance = item.category === '수당' && /교원수당|직원수당/.test(best.text);

  if (wrongBucket) {
    return {
      category: item.category,
      item: item.item,
      excelAmount: item.amount || null,
      pdfAmount,
      diff: numericDiff(item.amount, pdfAmount),
      status: '타 목 편성',
      reason: '정상 인건비 목이 아닌 다른 목에서 관련 편성 흔적을 확인했습니다.',
      evidence: pageSnippet(best),
    };
  }

  if (integratedAllowance) {
    return {
      category: item.category,
      item: item.item,
      excelAmount: item.amount || null,
      pdfAmount,
      diff: numericDiff(item.amount, pdfAmount),
      status: '통합편성 의심',
      reason: `${item.item}이 개별 표시되지 않고 교원수당 또는 직원수당에 통합 편성되었을 가능성이 있습니다.`,
      evidence: pageSnippet(best),
    };
  }

  if (!item.amount || !pdfAmount) {
    return {
      category: item.category,
      item: item.item,
      excelAmount: item.amount || null,
      pdfAmount: pdfAmount || null,
      diff: null,
      status: '확인필요',
      reason: '항목은 발견했으나 비교 가능한 금액을 안정적으로 추출하지 못했습니다.',
      evidence: pageSnippet(best),
    };
  }

  const diff = pdfAmount - item.amount;
  if (Math.abs(diff) <= 999) {
    return {
      category: item.category,
      item: item.item,
      excelAmount: item.amount,
      pdfAmount,
      diff: 0,
      status: '적정',
      reason: '엑셀 기준금액과 PDF 편성금액이 천원 단위 기준으로 일치합니다.',
      evidence: pageSnippet(best),
    };
  }

  return {
    category: item.category,
    item: item.item,
    excelAmount: item.amount,
    pdfAmount,
    diff,
    status: diff < 0 ? '과소편성' : '과다편성',
    reason: diff < 0 ? `PDF 편성금액이 엑셀 기준보다 ${Math.abs(diff).toLocaleString()}원 부족합니다.` : `PDF 편성금액이 엑셀 기준보다 ${diff.toLocaleString()}원 많습니다.`,
    evidence: pageSnippet(best),
  };
}

function reviewRetirement(retirement, pdfText, options) {
  const matches = findMatches(pdfText, ['퇴직적립금', '퇴직급여', '퇴직금', '퇴직'], HUMAN_COST_HINTS);
  if (matches.length > 0) {
    return {
      category: '퇴직',
      item: '퇴직적립금',
      excelAmount: '적립금액 있음',
      pdfAmount: '편성 있음',
      diff: '비교 제외',
      status: '적정',
      reason: '퇴직적립금 시트에 적립금액이 있고 세출예산명세서에도 퇴직 관련 예산 편성을 확인했습니다. 금액 비교는 제외했습니다.',
      evidence: pageSnippet(matches[0]),
    };
  }
  const wholePdfMatches = findMatches(pdfText, ['퇴직'], []);
  if (options.searchWholePdf && wholePdfMatches.length > 0) {
    return {
      category: '퇴직',
      item: '퇴직적립금',
      excelAmount: '적립금액 있음',
      pdfAmount: '다른 위치 의심',
      diff: '비교 제외',
      status: '확인필요',
      reason: '퇴직 관련 표현은 발견했으나 정상 편성 목인지 확인이 필요합니다.',
      evidence: pageSnippet(wholePdfMatches[0]),
    };
  }
  return {
    category: '퇴직',
    item: '퇴직적립금',
    excelAmount: '적립금액 있음',
    pdfAmount: '미편성',
    diff: '비교 제외',
    status: '부적정',
    reason: '퇴직적립금 시트에 적립금액 있음, 세출예산명세서에 미편성, 부적정',
    evidence: retirement.sourceRows?.[0] || '',
  };
}

function findMatches(pdfText, keywords, contextHints) {
  const normalizedKeywords = keywords.filter(Boolean).map(normalizeText);
  const matches = [];
  for (const page of pdfText.pages) {
    const sentences = splitContext(page.text);
    for (const sentence of sentences) {
      const normalized = normalizeText(sentence);
      const keywordScore = normalizedKeywords.reduce((score, kw) => score + (normalized.includes(kw) ? 3 : 0), 0);
      const contextScore = contextHints.reduce((score, kw) => score + (normalized.includes(kw) ? 1 : 0), 0);
      if (keywordScore > 0) {
        matches.push({ page: page.page, text: sentence, score: keywordScore + contextScore });
      }
    }
  }
  return matches.sort((a, b) => b.score - a.score);
}

function splitContext(text) {
  const clean = String(text || '').replace(/\s+/g, ' ');
  const chunks = [];
  for (let i = 0; i < clean.length; i += 260) chunks.push(clean.slice(i, i + 360));
  return chunks;
}

function inferPdfAmount(text, excelAmount, allowTruncation) {
  const formulaTotal = extractFormulaTotal(text);
  if (formulaTotal) return formulaTotal;
  const nums = extractNumbers(text).filter((n) => n > 0);
  if (!nums.length) return null;
  const candidates = new Set();
  for (const n of nums) {
    candidates.add(n);
    if (allowTruncation && n < 1000000) candidates.add(n * 1000);
  }
  if (excelAmount) {
    return [...candidates].sort((a, b) => Math.abs(a - excelAmount) - Math.abs(b - excelAmount))[0];
  }
  return Math.max(...candidates);
}

function extractFormulaTotal(text) {
  const clean = String(text || '').replace(/,/g, '');
  const eqMatches = [...clean.matchAll(/=\s*(\d{4,})/g)].map((m) => Number(m[1]));
  if (eqMatches.length) return Math.max(...eqMatches);
  return null;
}

function extractNumbers(text) {
  const clean = String(text || '').replace(/,/g, '');
  const matches = clean.match(/\d+(?:\.\d+)?/g) || [];
  return matches.map(Number).filter(Number.isFinite);
}

function numericDiff(a, b) {
  return typeof a === 'number' && typeof b === 'number' ? b - a : null;
}

function pageSnippet(match) {
  return `p.${match.page} ${String(match.text).replace(/\s+/g, ' ').slice(0, 180)}`;
}

function sourceEvidence(item) {
  return item.sourceRows?.[0] || '';
}

function prioritize(results) {
  const order = { '부적정': 0, '미편성': 1, '과소편성': 2, '과다편성': 3, '타 목 편성': 4, '통합편성 의심': 5, '확인필요': 6, '적정': 7 };
  return [...results].sort((a, b) => (order[a.status] ?? 99) - (order[b.status] ?? 99));
}

export function toCsv(rows) {
  const headers = ['대분류','항목','엑셀금액','PDF금액','차이','판정','상세사유','PDF 위치/근거'];
  const lines = [headers, ...rows.map((r) => [r.category, r.item, r.excelAmount, r.pdfAmount, r.diff, r.status, r.reason, r.evidence])];
  return '\ufeff' + lines.map((row) => row.map(csvEscape).join(',')).join('\n');
}

function csvEscape(value) {
  const text = value === null || value === undefined ? '' : String(value);
  return `"${text.replace(/"/g, '""')}"`;
}
