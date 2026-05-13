import { extractPdfText } from './pdf.js';
import { parseWorkbook, buildExcelFacts } from './excel.js';
import { runReview, toCsv } from './review.js';

const $ = (id) => document.getElementById(id);
let lastResults = [];

$('runBtn').addEventListener('click', async () => {
  const excelFile = $('excelFile').files?.[0];
  const pdfFile = $('pdfFile').files?.[0];
  if (!excelFile || !pdfFile) {
    alert('엑셀 파일과 PDF 파일을 모두 선택해주세요.');
    return;
  }
  $('runBtn').disabled = true;
  $('runBtn').textContent = '검토 중...';
  try {
    const workbook = await parseWorkbook(excelFile);
    const excelFacts = buildExcelFacts(workbook);
    const pdfText = await extractPdfText(pdfFile);
    const options = {
      allowTruncation: $('allowTruncation').checked,
      searchWholePdf: $('searchWholePdf').checked,
      retirementPresenceOnly: $('retirementPresenceOnly').checked,
    };
    lastResults = runReview(excelFacts, pdfText, options);
    renderSummary(lastResults, excelFacts, pdfText);
    renderTable(lastResults);
  } catch (error) {
    console.error(error);
    alert(`검토 중 오류가 발생했습니다: ${error.message}`);
  } finally {
    $('runBtn').disabled = false;
    $('runBtn').textContent = '검토 시작';
  }
});

$('downloadCsvBtn').addEventListener('click', () => {
  const blob = new Blob([toCsv(lastResults)], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `budget-review-${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

function renderSummary(results, excelFacts, pdfText) {
  const counts = results.reduce((acc, row) => {
    acc[row.status] = (acc[row.status] || 0) + 1;
    return acc;
  }, {});
  $('summary').hidden = false;
  $('summary').innerHTML = `
    <h2>요약</h2>
    <div class="summary-grid">
      <div class="metric"><strong>${results.length}</strong><span>검토 항목</span></div>
      <div class="metric"><strong>${counts['적정'] || 0}</strong><span>적정</span></div>
      <div class="metric"><strong>${(counts['과소편성'] || 0) + (counts['과다편성'] || 0)}</strong><span>금액 불일치</span></div>
      <div class="metric"><strong>${counts['미편성'] || 0}</strong><span>미편성</span></div>
      <div class="metric"><strong>${(counts['통합편성 의심'] || 0) + (counts['타 목 편성'] || 0) + (counts['부적정'] || 0)}</strong><span>확인 필요</span></div>
    </div>
    <p class="lead">엑셀 시트 ${excelFacts.sheetNames.length}개, PDF 텍스트 ${pdfText.raw.length.toLocaleString()}자를 분석했습니다. 숨겨진 세출예산명세서 보조근거 ${excelFacts.budgetPages?.length || 0}행도 함께 확인했습니다.</p>
    ${(excelFacts.diagnostics || []).length ? `<p class="warning">${excelFacts.diagnostics.map(escapeHtml).join('<br>')}</p>` : ''}`;
}

function renderTable(results) {
  $('resultsCard').hidden = false;
  const tbody = document.querySelector('#resultsTable tbody');
  tbody.innerHTML = '';
  for (const row of results) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(row.category)}</td>
      <td>${escapeHtml(row.item)}</td>
      <td class="amount">${formatValue(row.excelAmount)}</td>
      <td class="amount">${formatValue(row.pdfAmount)}</td>
      <td class="amount">${formatValue(row.diff)}</td>
      <td>${badge(row.status)}</td>
      <td>${escapeHtml(row.reason)}</td>
      <td>${escapeHtml(row.evidence || '')}</td>`;
    tbody.appendChild(tr);
  }
}

function formatValue(value) {
  if (value === null || value === undefined || value === '') return '-';
  if (typeof value === 'string') return escapeHtml(value);
  return `${Number(value).toLocaleString()}원`;
}

function badge(status) {
  const cls = status === '적정' ? 'ok' : ['과소편성','과다편성','미편성','부적정'].includes(status) ? 'bad' : status.includes('의심') || status.includes('타 목') ? 'warn' : 'info';
  return `<span class="badge ${cls}">${escapeHtml(status)}</span>`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"]/g, (ch) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[ch]));
}
