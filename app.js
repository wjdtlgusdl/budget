/* global XLSX */
import * as pdfjsLib from 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs';

const DEFAULTS = {
  teacherKeys: '교원,교사,원장,원감,담임,방과후,보조교사',
  staffKeys: '직원,일반직,행정,행정실장,조리,조리사,부조리사,운전,차량기사,사무,영양,청소',
  allowanceKeys: '연구수당,연구활동비,시간외수당,시간외근무수당,직급수당,직급보조비,관리업무수당,기타수당,자가운전보조금,교통보조비,급식보조비,식대보조비,정액급식비,식대,처우개선비,담임수당,교직수당,명절휴가비,명절동하계휴가비,휴일근무수당',
  salaryKeys: '본봉,기본급,급여,교원급여,직원급여,보수,봉급,원장급여,원감급여,차량기사급여,조리원급여',
  retireKeys: '퇴직,퇴직금,퇴직적립,퇴직적립금,퇴직금적립,퇴직급여,퇴직급여충당금,퇴직연금',
  bundleKeys: '교원수당,직원수당,수당,인건비'
};

const ALIASES = {
  '본봉': ['본봉','기본급','교원급여','직원급여','급여'],
  '교원급여': ['교원급여','본봉','기본급','급여'],
  '직원급여': ['직원급여','본봉','기본급','급여'],
  '연구수당': ['연구수당','연구활동비','연구비'],
  '연구활동비': ['연구수당','연구활동비','연구비'],
  '시간외수당': ['시간외수당','시간외근무수당'],
  '시간외근무수당': ['시간외수당','시간외근무수당'],
  '직급수당': ['직급수당','직급보조비'],
  '직급보조비': ['직급수당','직급보조비'],
  '교통보조비': ['교통보조비','자가운전보조금','자가운전보조비'],
  '자가운전보조금': ['교통보조비','자가운전보조금','자가운전보조비'],
  '식대보조비': ['식대보조비','급식보조비','정액급식비','교원정액급식비','직원정액급식비','식대'],
  '급식보조비': ['식대보조비','급식보조비','정액급식비','식대'],
  '명절휴가비': ['명절휴가비','명절동하계휴가비','명절동하계휴가'],
  '명절동하계휴가비': ['명절휴가비','명절동하계휴가비','명절동하계휴가'],
  '퇴직적립금': ['퇴직적립금','퇴직금적립','퇴직금및퇴직적립금','퇴직연금','퇴직급여']
};

let resultRows = [];
let showingOnlyIssues = false;

const $ = (id) => document.getElementById(id);
const splitKeys = (s) => String(s || '').split(',').map(v => v.trim()).filter(Boolean);
const norm = (v) => String(v ?? '').replace(/\s+/g, '').replace(/[()\[\]{}·ㆍ\-_/]/g, '').replace(/[0-9]+$/g,'').trim();
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
  const wb = XLSX.read(ab, { type:'array', cellDates:false, cellFormula:false, cellNF:false, cellText:false });
  return wb.SheetNames.map(name => ({
    name,
    rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, raw:true, defval:'' })
  }));
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

function cleanHeader(v){
  return String(v ?? '').replace(/\n/g,' ').replace(/\([A-Z=~0-9 ]+\)/gi,'').replace(/\s+/g,' ').trim();
}
function detectGroupFromSubtotal(text){
  const t = norm(text);
  if (t.includes('교원')) return '교원';
  if (t.includes('일반직') || t.includes('직원')) return '직원';
  return '';
}
function normalizeItemName(name, group){
  const n = norm(name);
  if (n === '본봉' || n === '기본급') return group === '직원' ? '직원급여' : '교원급여';
  return cleanHeader(name);
}
function itemCategory(item, cfg){
  if (hasAny(item, splitKeys(cfg.retireKeys))) return '퇴직적립금';
  if (hasAny(item, splitKeys(cfg.salaryKeys))) return '급여';
  if (hasAny(item, splitKeys(cfg.allowanceKeys))) return '수당';
  return '기타인건비';
}
function paymentColumns(rows, headerIdx){
  const headers = rows[headerIdx] || [];
  const cols = [];
  let started = false;
  for (let c=0; c<headers.length; c++) {
    const h = cleanHeader(headers[c]);
    const hn = norm(h);
    if (!h) continue;
    if (hn.includes('본봉') || hn.includes('기본급')) started = true;
    if (!started) continue;
    if (/지급액계|소득세|주민세|본인부담금|공제액계|실수령액/.test(hn)) break;
    if (!/일련|직명|성명|경력|호봉|주민/.test(hn)) cols.push({idx:c, label:h});
  }
  return cols;
}

function extractPayrollBySubtotal(sheet, cfg){
  const items = [];
  const headerIdx = sheet.rows.findIndex(r => r.some(v => norm(v).includes('본봉')) && r.some(v => norm(v).includes('지급액계')));
  if (headerIdx < 0) return items;
  const cols = paymentColumns(sheet.rows, headerIdx);
  for (let r=headerIdx+1; r<sheet.rows.length; r++) {
    const row = sheet.rows[r];
    const rowText = row.join(' ');
    if (!norm(rowText).includes('소계')) continue;
    const group = detectGroupFromSubtotal(rowText);
    if (!group) continue;
    for (const col of cols) {
      const amount = money(row[col.idx]);
      if (amount === null || amount <= 0) continue;
      const item = normalizeItemName(col.label, group);
      items.push({
        group,
        category: itemCategory(item, cfg),
        item,
        displayItem: col.label,
        amount,
        sheet: sheet.name,
        sourceRow: r+1,
        rawText: rowText,
        matchTerms: aliasesFor(item, group)
      });
    }
  }
  return items;
}

function extractGenericRetirement(sheets, cfg){
  const retireKeys = splitKeys(cfg.retireKeys);
  const out = [];
  for (const sheet of sheets) {
    if (!hasAny(sheet.name, retireKeys)) continue;
    for (let i=0; i<sheet.rows.length; i++) {
      const row = sheet.rows[i];
      const text = row.join(' ');
      const compact = norm(text);
      if (!compact || /해당없음|해당사항없음|공란|없음/.test(compact)) continue;
      const nums = row.map(money).filter(v => v !== null && v > 0);
      if (!nums.length) continue;
      const amount = nums.reduce((a,b)=> Math.abs(b) > Math.abs(a) ? b : a, nums[0]);
      out.push({ group:'미분류', category:'퇴직적립금', item:'퇴직적립금', displayItem:'퇴직적립금', amount, sheet:sheet.name, sourceRow:i+1, rawText:text, matchTerms:aliasesFor('퇴직적립금','') });
    }
  }
  return out;
}

function extractBudgetItems(sheets, cfg){
  let items = [];
  for (const sheet of sheets) {
    if (hasAny(sheet.name, ['보수일람표','교직원보수'])) {
      items = items.concat(extractPayrollBySubtotal(sheet, cfg));
    }
  }
  items = items.concat(extractGenericRetirement(sheets, cfg));
  return dedupeItems(items);
}
function dedupeItems(items){
  const map = new Map();
  for (const it of items) {
    const key = [it.group,it.category,norm(it.item),it.amount,it.sheet,it.sourceRow].join('|');
    if (!map.has(key)) map.set(key,it);
  }
  return [...map.values()];
}
function aliasesFor(item, group){
  const base = ALIASES[item] || ALIASES[cleanHeader(item)] || [item];
  if (item === '교원급여') return ['교원급여'];
  if (item === '직원급여') return ['직원급여'];
  if (item === '퇴직적립금') return group === '교원' ? ['교원퇴직금및퇴직적립금','퇴직연금','퇴직적립금'] : ['직원퇴직금및퇴직적립금','퇴직연금','퇴직적립금'];
  return [...new Set(base)];
}

function buildNormMap(text){
  let compact = '', map = [];
  for (let i=0; i<text.length; i++) {
    const ch = text[i];
    if (/\s/.test(ch) || /[()\[\]{}·ㆍ\-_/]/.test(ch)) continue;
    compact += ch;
    map.push(i);
  }
  return { compact: compact.replace(/[0-9]+(?=[^0-9]|$)/g, m => m), map };
}
function amountCandidates(window){
  const out = [];
  const afterEquals = [...window.matchAll(/=\s*([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{5,})/g)].map(m => money(m[1])).filter(v => v !== null);
  afterEquals.forEach(v => out.push({value:v, source:'산출기초'}));
  const nums = [...window.matchAll(/([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{5,})/g)].map(m => money(m[1])).filter(v => v !== null);
  nums.forEach(v => out.push({value:v, source:'주변숫자'}));
  return out;
}
function immediateSectionTotal(window){
  // 목 예산 행: 항목명 283,556 216,000 0 499,556 형태는 단위가 천원입니다.
  const beforePlan = window.split(/\(본예산\)|발행일|\n/)[0];
  const nums = [...beforePlan.matchAll(/\b([0-9]{1,3}(?:,[0-9]{3})+|0)\b/g)].map(m => money(m[1])).filter(v => v !== null);
  if (nums.length >= 4) return nums[3] * 1000;
  if (nums.length >= 3) return nums[nums.length-1] * 1000;
  return null;
}
function extractStatementAmount(statementText, terms, options={}){
  const { compact, map } = buildNormMap(statementText);
  const found = [];
  for (const term of terms) {
    const nt = norm(term);
    if (!nt) continue;
    let pos = 0;
    while ((pos = compact.indexOf(nt, pos)) !== -1) {
      const orig = map[pos] ?? 0;
      const window = statementText.slice(Math.max(0, orig - 80), Math.min(statementText.length, orig + 360));
      const section = immediateSectionTotal(window);
      const candidates = amountCandidates(window);
      found.push({term, pos:orig, window, section, candidates});
      pos += nt.length;
    }
  }
  if (!found.length) return {found:false, amount:null, term:'', evidence:''};

  if (options.preferSectionTotal) {
    const withSection = found.find(f => f.section !== null && f.section > 0);
    if (withSection) return {found:true, amount:withSection.section, term:withSection.term, evidence:withSection.window.trim().slice(0,240)};
  }
  const exactTerm = found.find(f => f.candidates.some(c => c.source === '산출기초'));
  if (exactTerm) return {found:true, amount:exactTerm.candidates.find(c => c.source === '산출기초').value, term:exactTerm.term, evidence:exactTerm.window.trim().slice(0,240)};
  const anySection = found.find(f => f.section !== null && f.section > 0);
  if (anySection) return {found:true, amount:anySection.section, term:anySection.term, evidence:anySection.window.trim().slice(0,240)};
  const any = found.find(f => f.candidates.length);
  return {found:true, amount:any ? any.candidates[0].value : null, term: found[0].term, evidence: found[0].window.trim().slice(0,240)};
}
function bundleFound(statementText, group, cfg){
  const b = splitKeys(cfg.bundleKeys).flatMap(k => k === '수당' && group && group !== '미분류' ? [`${group}${k}`, k] : [k]);
  const compact = norm(statementText);
  return b.some(k => compact.includes(norm(k)));
}
function compareItems(items, statementText, cfg){
  const rows = items.map(it => {
    const preferSectionTotal = /교원급여|직원급여|퇴직/.test(norm(it.item));
    const found = extractStatementAmount(statementText, it.matchTerms || aliasesFor(it.item,it.group), {preferSectionTotal});
    const statementAmount = found.found ? found.amount : null;
    const diff = statementAmount === null ? null : it.amount - statementAmount;
    const hasBundle = it.category === '수당' && bundleFound(statementText, it.group, cfg);
    let status, memo;
    if (found.found && statementAmount !== null && diff === 0) { status = '일치'; memo = '항목 및 금액 일치'; }
    else if (found.found && statementAmount !== null && diff > 0) { status = '과소편성'; memo = `세출예산명세서가 ${fmt(Math.abs(diff))}원 적게 편성되어 있습니다.`; }
    else if (found.found && statementAmount !== null && diff < 0) { status = '과다편성'; memo = `세출예산명세서가 ${fmt(Math.abs(diff))}원 많이 편성되어 있습니다.`; }
    else if (!found.found && hasBundle) { status = '개별 미편성'; memo = `${it.displayItem || it.item} 개별 산출내역은 없고 ${it.group}수당 등 통합 편성 가능성이 있습니다.`; }
    else { status = '명세서 내역 없음'; memo = '세출예산명세서에서 대응 항목을 찾지 못했습니다.'; }

    return {
      구분: it.group,
      검토분류: it.category,
      예산서항목: it.displayItem || it.item,
      비교기준항목: it.item,
      예산서금액: it.amount,
      명세서대응항목: found.found ? found.term : (hasBundle ? `${it.group}수당 통합 항목` : ''),
      명세서금액: statementAmount,
      차액: diff,
      판정: status,
      검토메모: memo,
      근거일부: found.evidence || '',
      예산서시트: it.sheet,
      예산서행: it.sourceRow
    };
  });
  return rows.concat(statementOnlyChecks(items, statementText));
}


function groupSegment(statementText, group){
  const { compact, map } = buildNormMap(statementText);
  const startKey = group === '직원' ? '직원인건비' : group === '교원' ? '교원인건비' : '';
  if (!startKey) return statementText;
  const start = compact.indexOf(norm(startKey));
  if (start < 0) return statementText;
  const endKeys = group === '교원' ? ['직원인건비','그밖의인건비','운영비'] : ['그밖의인건비','운영비','관리운영비'];
  let end = compact.length;
  for (const k of endKeys) {
    const e = compact.indexOf(norm(k), start + norm(startKey).length);
    if (e > -1 && e < end) end = e;
  }
  return statementText.slice(map[start] ?? 0, map[end] ?? statementText.length);
}

function statementOnlyChecks(items, statementText){
  const checks = [
    {group:'교원', category:'수당', item:'명절휴가비', terms:['명절휴가비','명절동하계휴가비','명절동하계휴가']},
    {group:'직원', category:'수당', item:'명절휴가비', terms:['명절휴가비','명절동하계휴가비','명절동하계휴가']}
  ];
  const rows = [];
  for (const chk of checks) {
    const existsInBudget = items.some(it => it.group === chk.group && chk.terms.some(t => norm(it.item).includes(norm(t)) || norm(it.displayItem).includes(norm(t))));
    if (existsInBudget) continue;
    const found = extractStatementAmount(groupSegment(statementText, chk.group), chk.terms, {preferSectionTotal:false});
    if (!found.found || found.amount === null) continue;
    // Avoid duplicating one global occurrence twice unless both groups have their own 수당 area; still useful as warning.
    rows.push({
      구분: chk.group,
      검토분류: chk.category,
      예산서항목: '',
      비교기준항목: chk.item,
      예산서금액: null,
      명세서대응항목: found.term,
      명세서금액: found.amount,
      차액: null,
      판정: '보수일람표 미반영',
      검토메모: `${chk.item}가 세출예산명세서 산출내역에 있으나 교직원보수일람표 소계 항목에는 없습니다.`,
      근거일부: found.evidence || '',
      예산서시트: '',
      예산서행: ''
    });
  }
  return rows;
}

function renderSummary(rows){
  const el = $('summary');
  const total = rows.length;
  const issues = rows.filter(r => r.판정 !== '일치').length;
  const diff = rows.filter(r => /과소편성|과다편성|금액/.test(r.판정)).length;
  const missing = rows.filter(r => /미편성|내역 없음|미반영/.test(r.판정)).length;
  el.hidden = false;
  el.innerHTML = [
    ['전체 검토', total], ['문제 항목', issues], ['금액 차이', diff], ['미편성/누락', missing]
  ].map(([a,b]) => `<div class="kpi"><span>${a}</span><b>${b}</b></div>`).join('');
}

function renderTable(){
  const rows = showingOnlyIssues ? resultRows.filter(r => r.판정 !== '일치') : resultRows;
  const headers = ['구분','검토분류','예산서항목','비교기준항목','예산서금액','명세서대응항목','명세서금액','차액','판정','검토메모','예산서시트','예산서행'];
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
    if (!items.length) alert('교직원보수일람표의 소계 행을 찾지 못했습니다. 예산서 양식의 헤더/소계 행을 확인해주세요.');
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
  const ws = XLSX.utils.json_to_sheet(resultRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '인건비_대조결과');
  XLSX.writeFile(wb, `인건비_대조결과_${new Date().toISOString().slice(0,10)}.xlsx`);
};
