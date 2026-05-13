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
  '비정기수당': ['비정기수당','명절휴가비','행사수당'],
  '관리업무수당': ['관리업무수당'],
  '직책수당': ['직책수당'],
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
      text += `\n[[PAGE ${p}]]\n` + pdfItemsToLines(content.items).join('\n');
    }
    return text;
  }
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    const sheets = await workbookToRows(file);
    return sheets.map(s => `\n[${s.name}]\n` + s.rows.map(r => r.join(' ')).join('\n')).join('\n');
  }
  return await readTextFile(file);
}

function pdfItemsToLines(items){
  const rows = [];
  for (const it of items) {
    const str = String(it.str || '').trim();
    if (!str) continue;
    const tx = it.transform || [1,0,0,1,0,0];
    const x = tx[4] || 0;
    const y = Math.round((tx[5] || 0) * 2) / 2;
    let row = rows.find(r => Math.abs(r.y - y) <= 2);
    if (!row) { row = {y, items:[]}; rows.push(row); }
    row.items.push({x, str});
  }
  rows.sort((a,b) => b.y - a.y);
  return rows.map(r => r.items.sort((a,b)=>a.x-b.x).map(i=>i.str).join(' ').replace(/\s+/g,' ').trim()).filter(Boolean);
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
  return { compact, map };
}
function lineInfo(statementText){
  return statementText.split(/\n+/).map(x => x.replace(/\s+/g,' ').trim()).filter(Boolean);
}
function parseLineAmounts(line){
  return [...String(line).matchAll(/([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{5,}|0)/g)].map(m => money(m[1])).filter(v => v !== null);
}
function parseBudgetRow(line, nextLine='', next2Line=''){
  const candidates = [line];
  if (!/[\[\]원*=]|본예산/.test(line)) {
    candidates.push(`${line} ${nextLine}`.replace(/\s+/g,' '));
    candidates.push(`${line} ${nextLine} ${next2Line}`.replace(/\s+/g,' '));
  }
  for (const raw of candidates) {
    if (!raw || /\(본예산\)|원\s*\*|=|발행일|예산구분|과\s*목|산출내역|산\s*출\s*기\s*초/.test(raw)) continue;
    const m = raw.match(/^(.{1,45}?)\s+([0-9]{1,3}(?:,[0-9]{3})+|0)\s+([0-9]{1,3}(?:,[0-9]{3})+|0)\s+([0-9]{1,3}(?:,[0-9]{3})+|0)\s+([0-9]{1,3}(?:,[0-9]{3})+|0)(?:\s|$)/);
    if (!m) continue;
    const name = m[1].replace(/\s+/g,' ').trim();
    if (!name || /단\s*위|보조금|수익자|합계\s*산출/.test(name)) continue;
    return {
      name,
      support: money(m[2]) * 1000,
      parent: money(m[3]) * 1000,
      other: money(m[4]) * 1000,
      total: money(m[5]) * 1000
    };
  }
  return null;
}
function isMajorOrSubRow(name){
  const n = norm(name);
  return ['인건비','교원인건비','직원인건비','그밖의인건비'].includes(n) || n.endsWith('인건비');
}
function groupFromRowName(name, currentGroup){
  const n = norm(name);
  if (n.includes('교원인건비')) return '교원';
  if (n.includes('직원인건비')) return '직원';
  if (n.includes('그밖의인건비')) return '그밖의';
  return currentGroup;
}
function parseStatementEntries(statementText){
  const lines = lineInfo(statementText);
  const entries = [];
  let currentGroup = '';
  let currentEntry = null;
  let inExpense = false;
  for (let i=0; i<lines.length; i++) {
    const line = lines[i];
    if (norm(line).includes('세출예산명세서')) inExpense = true;
    if (!inExpense) continue;
    const row = parseBudgetRow(line, lines[i+1] || '', lines[i+2] || '');
    if (row) {
      const rowGroup = groupFromRowName(row.name, currentGroup);
      const n = norm(row.name);
      if (n === '인건비' || n === '운영비' || n.endsWith('활동비') && !currentGroup) {
        currentEntry = null;
      }
      if (n.includes('교원인건비') || n.includes('직원인건비') || n.includes('그밖의인건비')) {
        currentGroup = rowGroup;
        currentEntry = null;
      } else if (currentGroup && !isMajorOrSubRow(row.name)) {
        currentEntry = { group: currentGroup, item: row.name, sectionTotal: row.total, support: row.support, parent: row.parent, other: row.other, basisLines: [], startLine: i };
        entries.push(currentEntry);
      } else {
        currentEntry = null;
      }
      continue;
    }
    if (currentEntry) currentEntry.basisLines.push(line);
  }
  return entries;
}
function nearbyBudgetLabel(lines, i){
  const isLabel = (s) => /\[[^\]]+\].+/.test(s) && !/=/.test(s);
  for (let j=i+1; j<=Math.min(lines.length-1, i+2); j++) if (isLabel(lines[j])) return lines[j];
  if (/\[[^\]]+\]/.test(lines[i])) return lines[i];
  for (let j=i-1; j>=Math.max(0, i-2); j--) if (isLabel(lines[j])) return lines[j];
  return '';
}
function allFormulaAmounts(sectionLines){
  const amounts = [];
  for (const line of sectionLines) {
    [...line.matchAll(/=\s*([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{5,})/g)]
      .map(m => money(m[1])).filter(v => v !== null).forEach(v => amounts.push(v));
  }
  return amounts;
}
function basisAmountsFromLines(sectionLines, terms){
  const nterms = terms.map(norm).filter(Boolean);
  const amounts = [];
  for (let i=0; i<sectionLines.length; i++) {
    const line = sectionLines[i];
    const eqs = [...line.matchAll(/=\s*([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{5,})/g)].map(m => money(m[1])).filter(v => v !== null);
    if (!eqs.length) continue;
    const context = `${line} ${nearbyBudgetLabel(sectionLines, i)}`;
    const nctx = norm(context);
    if (nterms.some(t => nctx.includes(t))) eqs.forEach(v => amounts.push(v));
  }
  return amounts;
}
function entryEvidence(entry, terms){
  const nterms = terms.map(norm).filter(Boolean);
  const lines = [];
  for (let i=0; i<entry.basisLines.length; i++) {
    const line = entry.basisLines[i];
    if (!/=/.test(line)) continue;
    const ctx = `${line} ${nearbyBudgetLabel(entry.basisLines, i)}`;
    if (nterms.some(t => norm(ctx).includes(t))) lines.push(ctx);
  }
  const head = `${entry.group}인건비 > ${entry.item} 목합계 ${fmt(entry.sectionTotal)}원`;
  return [head, ...lines.slice(0,4)].join(' / ').slice(0,500);
}
function entryMatchesItem(entry, terms, group){
  if (group && entry.group !== group) return false;
  const nterms = terms.map(norm).filter(Boolean);
  const itemName = norm(entry.item);
  if (nterms.some(t => itemName.includes(t) || t.includes(itemName))) return true;
  return basisAmountsFromLines(entry.basisLines, terms).length > 0;
}
function extractStatementEntry(statementText, terms, group){
  const entries = parseStatementEntries(statementText);
  const scoped = entries.filter(e => !group || e.group === group);
  const direct = scoped.find(e => terms.some(t => norm(e.item).includes(norm(t)) || norm(t).includes(norm(e.item))));
  if (direct) {
    const basisAmounts = allFormulaAmounts(direct.basisLines);
    const basisTotal = basisAmounts.length ? basisAmounts.reduce((a,b)=>a+b,0) : null;
    return {found:true, term:direct.item, sectionTotal:direct.sectionTotal, basisTotal, amount:direct.sectionTotal ?? basisTotal, evidence:entryEvidence(direct, terms)};
  }
  const basisMatches = scoped.map(e => ({entry:e, amounts:basisAmountsFromLines(e.basisLines, terms)})).filter(x => x.amounts.length);
  if (basisMatches.length) {
    const e = basisMatches[0].entry;
    const basisTotal = basisMatches.reduce((sum,x)=>sum+x.amounts.reduce((a,b)=>a+b,0),0);
    return {found:true, term:e.item, sectionTotal:e.sectionTotal, basisTotal, amount:basisTotal, evidence:entryEvidence(e, terms)};
  }
  return {found:false, term:'', sectionTotal:null, basisTotal:null, amount:null, evidence:''};
}
function extractStatementAmount(statementText, terms, options={}){
  const entry = extractStatementEntry(statementText, terms, options.group || '');
  return {found:entry.found, amount:entry.amount, term:entry.term, evidence:entry.evidence, sectionTotal:entry.sectionTotal, basisTotal:entry.basisTotal};
}
function bundleFound(statementText, group, cfg){
  const b = splitKeys(cfg.bundleKeys).flatMap(k => k === '수당' && group && group !== '미분류' ? [`${group}${k}`, k] : [k]);
  const compact = norm(statementText);
  return b.some(k => compact.includes(norm(k)));
}
function compareItems(items, statementText, cfg){
  const rows = items.map(it => {
    const isRetirement = it.category === '퇴직적립금';
    const preferSectionTotal = /교원급여|직원급여/.test(norm(it.item));
    const found = extractStatementAmount(statementText, it.matchTerms || aliasesFor(it.item,it.group), {preferSectionTotal, group: it.group});
    const statementAmount = (!isRetirement && found.found) ? found.amount : null;
    const sectionTotal = (!isRetirement && found.found) ? found.sectionTotal : null;
    const basisTotal = (!isRetirement && found.found) ? found.basisTotal : null;
    const diff = (!isRetirement && statementAmount !== null) ? it.amount - statementAmount : null;
    const hasBundle = it.category === '수당' && bundleFound(statementText, it.group, cfg);
    let status, memo;

    if (isRetirement) {
      if (found.found) {
        status = '항목 있음';
        memo = '퇴직적립금은 금액 비교 제외 대상입니다. 세출예산명세서에 퇴직적립금 관련 항목 존재 여부만 확인했습니다.';
      } else {
        status = '명세서 항목 없음';
        memo = '엑셀의 퇴직적립금/퇴직금적립 관련 시트에 금액이 있으나 세출예산명세서에서 퇴직적립금 관련 항목을 찾지 못했습니다.';
      }
    } else if (found.found && statementAmount !== null && diff === 0) { status = '일치'; memo = '항목 및 금액 일치'; }
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
      명세서목합계: sectionTotal,
      산출기초합계: basisTotal,
      비교명세서금액: statementAmount,
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
    const found = extractStatementAmount(statementText, chk.terms, {group: chk.group});
    if (!found.found || found.amount === null) continue;
    // Avoid duplicating one global occurrence twice unless both groups have their own 수당 area; still useful as warning.
    rows.push({
      구분: chk.group,
      검토분류: chk.category,
      예산서항목: '',
      비교기준항목: chk.item,
      예산서금액: null,
      명세서대응항목: found.term,
      명세서목합계: found.sectionTotal,
      산출기초합계: found.basisTotal,
      비교명세서금액: found.amount,
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
  const issues = rows.filter(r => !['일치','항목 있음'].includes(r.판정)).length;
  const diff = rows.filter(r => /과소편성|과다편성|금액/.test(r.판정)).length;
  const missing = rows.filter(r => /미편성|내역 없음|미반영|항목 없음/.test(r.판정)).length;
  el.hidden = false;
  el.innerHTML = [
    ['전체 검토', total], ['문제 항목', issues], ['금액 차이', diff], ['미편성/누락', missing]
  ].map(([a,b]) => `<div class="kpi"><span>${a}</span><b>${b}</b></div>`).join('');
}

function renderTable(){
  const rows = showingOnlyIssues ? resultRows.filter(r => !['일치','항목 있음'].includes(r.판정)) : resultRows;
  const headers = ['구분','검토분류','예산서항목','비교기준항목','예산서금액','명세서대응항목','명세서목합계','산출기초합계','비교명세서금액','차액','판정','검토메모','예산서시트','예산서행'];
  const table = $('resultTable');
  table.innerHTML = '<thead><tr>' + headers.map(h=>`<th>${h}</th>`).join('') + '</tr></thead><tbody>' +
    rows.map(r => '<tr>' + headers.map(h => {
      const val = ['예산서금액','명세서목합계','산출기초합계','비교명세서금액','차액'].includes(h) ? fmt(r[h]) : (r[h] ?? '');
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
