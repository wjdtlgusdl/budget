/* global XLSX */
import * as pdfjsLib from 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs';

const DEFAULTS = {
  teacherKeys: '교원,교사,원장,원감,담임,보조교사',
  staffKeys: '직원,행정,행정실장,조리,조리사,운전,사무,영양',
  allowanceKeys: '연구수당,연구활동비,시간외수당,직급보조비,관리업무수당,기타수당,자가운전보조금,정액급식비,식대,처우개선비,담임수당,교직수당',
  salaryKeys: '본봉,기본급,급여,보수,봉급',
  retireKeys: '퇴직,퇴직금,퇴직적립,퇴직적립금,퇴직금적립,퇴직급여,퇴직급여충당금',
  bundleKeys: '교원수당,직원수당,수당,인건비'
};

let resultRows = [];
let showingOnlyIssues = false;

const $ = (id) => document.getElementById(id);
const splitKeys = (s) => String(s || '').split(',').map(v => v.trim()).filter(Boolean);
const norm = (v) => String(v ?? '').replace(/\s+/g, '').replace(/[()\[\]{}]/g, '').trim();
const hasAny = (text, keys) => keys.some(k => norm(text).includes(norm(k)));
const money = (v) => {
  if (typeof v === 'number') return Math.round(v);
  const s = String(v ?? '').replace(/[^0-9.-]/g, '');
  if (!s || s === '-' || s === '.') return null;
  const n = Number(s);
  return Number.isFinite(n) ? Math.round(n) : null;
};
const fmt = (n) => n === null || n === undefined || Number.isNaN(n) ? '' : Number(n).toLocaleString('ko-KR');

function initConfig(){ Object.entries(DEFAULTS).forEach(([id,val]) => $(id).value = val); }
initConfig();
$('resetConfig').onclick = initConfig;

async function readArrayBuffer(file){ return await file.arrayBuffer(); }
async function readTextFile(file){ return await file.text(); }

async function workbookToRows(file){
  const ab = await readArrayBuffer(file);
  const wb = XLSX.read(ab, { type:'array', cellDates:false });
  const sheets = [];
  wb.SheetNames.forEach(name => {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, raw:false, defval:'' });
    sheets.push({ name, rows: aoa });
  });
  return sheets;
}

async function fileToStatementText(file){
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.pdf')) {
    const pdf = await pdfjsLib.getDocument({ data: await readArrayBuffer(file) }).promise;
    let text = '';
    for (let p=1; p<=pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      text += '\n' + content.items.map(i => i.str).join(' ');
    }
    return text;
  }
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    const sheets = await workbookToRows(file);
    return sheets.map(s => `\n[${s.name}]\n` + s.rows.map(r => r.join(' ')).join('\n')).join('\n');
  }
  return await readTextFile(file);
}

function detectGroup(rowText, activeGroup, teacherKeys, staffKeys){
  if (hasAny(rowText, staffKeys)) return '직원';
  if (hasAny(rowText, teacherKeys)) return '교원';
  return activeGroup || '미분류';
}

function extractBudgetItems(sheets, cfg){
  const teacherKeys = splitKeys(cfg.teacherKeys);
  const staffKeys = splitKeys(cfg.staffKeys);
  const allowanceKeys = splitKeys(cfg.allowanceKeys);
  const salaryKeys = splitKeys(cfg.salaryKeys);
  const retireKeys = splitKeys(cfg.retireKeys);
  const allPayKeys = [...allowanceKeys, ...salaryKeys];
  const items = [];

  for (const sheet of sheets) {
    let activeGroup = '';
    const isRetireSheet = hasAny(sheet.name, retireKeys);
    for (let i=0; i<sheet.rows.length; i++) {
      const row = sheet.rows[i];
      const text = row.join(' ');
      const compact = norm(text);
      if (!compact || /해당없음|해당사항없음|공란|없음|총칙/.test(compact)) continue;
      activeGroup = detectGroup(text, activeGroup, teacherKeys, staffKeys);
      const nums = row.map(money).filter(v => v !== null && Math.abs(v) > 0);
      if (!nums.length) continue;
      const amount = nums.reduce((a,b)=> Math.abs(b) > Math.abs(a) ? b : a, nums[0]);

      if (isRetireSheet || hasAny(text, retireKeys)) {
        items.push({ group: activeGroup, category:'퇴직적립금', item: bestKeyword(text, retireKeys) || '퇴직적립금', amount, sheet: sheet.name, sourceRow: i+1, rawText: text });
        continue;
      }

      const key = bestKeyword(text, allPayKeys);
      if (key) {
        const category = hasAny(key, allowanceKeys) ? '수당' : '급여';
        items.push({ group: activeGroup, category, item: key, amount, sheet: sheet.name, sourceRow: i+1, rawText: text });
      }
    }
  }
  return dedupeItems(items);
}

function bestKeyword(text, keys){
  const compact = norm(text);
  const sorted = [...keys].sort((a,b)=>norm(b).length-norm(a).length);
  return sorted.find(k => compact.includes(norm(k))) || '';
}

function dedupeItems(items){
  const map = new Map();
  for (const it of items) {
    const key = [it.group,it.category,norm(it.item),it.amount,it.sheet].join('|');
    if (!map.has(key)) map.set(key,it);
  }
  return [...map.values()];
}

function extractNearbyAmounts(statementText, keyword){
  const clean = statementText.replace(/\s+/g, ' ');
  const idxs = [];
  let pos = 0;
  const nk = norm(keyword);
  const nclean = norm(clean);
  while ((pos = nclean.indexOf(nk, pos)) !== -1) { idxs.push(pos); pos += nk.length; }
  const amounts = [];
  // use original text windows by searching non-normalized keyword too, with fallback global Korean words removed is not perfect.
  const re = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
  let m;
  while ((m = re.exec(clean)) !== null) {
    const window = clean.slice(Math.max(0, m.index - 80), Math.min(clean.length, m.index + 160));
    const found = window.match(/[-+]?\d{1,3}(?:,\d{3})+|[-+]?\d{5,}/g) || [];
    found.map(money).filter(v => v !== null).forEach(v => amounts.push(v));
  }
  return [...new Set(amounts)];
}

function compareItems(items, statementText, cfg){
  const bundleKeys = splitKeys(cfg.bundleKeys);
  const statementCompact = norm(statementText);
  return items.map(it => {
    const itemFound = statementCompact.includes(norm(it.item));
    const amounts = itemFound ? extractNearbyAmounts(statementText, it.item) : [];
    let statementAmount = amounts.find(a => a === it.amount);
    if (statementAmount === undefined) statementAmount = amounts.length ? amounts[0] : null;
    const diff = statementAmount === null ? null : it.amount - statementAmount;
    const bundleFound = bundleKeys.some(k => statementCompact.includes(norm(groupBundle(it.group,k)) ) || statementCompact.includes(norm(k)));

    let status = '명세서 내역 없음';
    let memo = '세출예산명세서에서 대응 항목을 찾지 못했습니다.';
    if (itemFound && statementAmount !== null && diff === 0) { status = '일치'; memo = '항목 및 금액 일치'; }
    else if (itemFound && statementAmount !== null && diff !== 0) { status = '금액 불일치'; memo = `${fmt(Math.abs(diff))}원 차이`; }
    else if (it.category === '수당' && bundleFound) { status = '개별 미편성'; memo = `${it.item} 개별 항목은 없고 교원수당/직원수당 등 통합 편성 가능성이 있습니다.`; }
    else if (it.category === '퇴직적립금') { status = '명세서 내역 없음'; memo = '퇴직적립금/퇴직금적립 관련 금액이 예산서에 있으나 명세서 대응 내역이 확인되지 않습니다.'; }

    return {
      구분: it.group,
      검토분류: it.category,
      예산서항목: it.item,
      예산서금액: it.amount,
      명세서대응항목: itemFound ? it.item : (it.category === '수당' && bundleFound ? '통합 수당 항목 가능' : ''),
      명세서금액: statementAmount,
      차액: diff,
      판정: status,
      검토메모: memo,
      예산서시트: it.sheet,
      예산서행: it.sourceRow
    };
  });
}
function groupBundle(group, key){ return group && group !== '미분류' && key === '수당' ? `${group}${key}` : key; }

function renderSummary(rows){
  const el = $('summary');
  const total = rows.length;
  const issues = rows.filter(r => r.판정 !== '일치').length;
  const diff = rows.filter(r => r.판정 === '금액 불일치').length;
  const missing = rows.filter(r => /미편성|내역 없음/.test(r.판정)).length;
  el.hidden = false;
  el.innerHTML = [
    ['전체 검토', total], ['문제 항목', issues], ['금액 차이', diff], ['미편성/누락', missing]
  ].map(([a,b]) => `<div class="kpi"><span>${a}</span><b>${b}</b></div>`).join('');
}

function renderTable(){
  const rows = showingOnlyIssues ? resultRows.filter(r => r.판정 !== '일치') : resultRows;
  const headers = ['구분','검토분류','예산서항목','예산서금액','명세서대응항목','명세서금액','차액','판정','검토메모','예산서시트','예산서행'];
  const table = $('resultTable');
  table.innerHTML = '<thead><tr>' + headers.map(h=>`<th>${h}</th>`).join('') + '</tr></thead><tbody>' +
    rows.map(r => '<tr>' + headers.map(h => {
      const val = ['예산서금액','명세서금액','차액'].includes(h) ? fmt(r[h]) : (r[h] ?? '');
      const cls = h === '판정' ? ` class="status-${String(r[h]).replace(/\s/g,'')}"` : '';
      return `<td${cls}>${escapeHtml(val)}</td>`;
    }).join('') + '</tr>').join('') + '</tbody>';
  $('resultCard').hidden = false;
}
function escapeHtml(s){ return String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

$('runBtn').onclick = async () => {
  const budgetFile = $('budgetFile').files[0];
  const statementFile = $('statementFile').files[0];
  if (!budgetFile || !statementFile) { alert('예산서 엑셀 파일과 세출예산명세서 파일을 모두 선택해주세요.'); return; }
  $('runBtn').disabled = true;
  $('runBtn').textContent = '분석 중...';
  try {
    const cfg = Object.fromEntries(Object.keys(DEFAULTS).map(k => [k, $(k).value]));
    const sheets = await workbookToRows(budgetFile);
    const statementText = await fileToStatementText(statementFile);
    const items = extractBudgetItems(sheets, cfg);
    resultRows = compareItems(items, statementText, cfg);
    renderSummary(resultRows);
    renderTable();
  } catch (e) {
    console.error(e);
    alert('분석 중 오류가 발생했습니다. 파일 형식이나 브라우저 콘솔을 확인해주세요.');
  } finally {
    $('runBtn').disabled = false;
    $('runBtn').textContent = '대조 실행';
  }
};

$('onlyIssue').onclick = () => { showingOnlyIssues = !showingOnlyIssues; $('onlyIssue').textContent = showingOnlyIssues ? '전체 보기' : '문제 항목만 보기'; renderTable(); };
$('downloadXlsx').onclick = () => {
  if (!resultRows.length) return;
  const ws = XLSX.utils.json_to_sheet(resultRows.map(r => ({...r, 예산서금액: r.예산서금액, 명세서금액: r.명세서금액, 차액: r.차액})));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '인건비_대조결과');
  XLSX.writeFile(wb, `인건비_대조결과_${new Date().toISOString().slice(0,10)}.xlsx`);
};
