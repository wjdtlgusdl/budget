/* global XLSX, pdfjsLib, JSZip */
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const $ = (id) => document.getElementById(id);
const state = { report: null };

const MONEY_KEYS = ['금액','합계','계','지급액계','본봉','기본급','수당','급여','보조금'];
const LEGAL_RE = /법정부담|4대보험|사학연금|국민연금|건강보험|고용보험|산재보험/;
const SALARY_SHEET_RE = /(교직원.*보수|보수.*일람|보수|급여|봉급|인건비|교직원)/;
const RETIRE_RE = /퇴직/;

function norm(v){ return String(v ?? '').replace(/\s+/g,'').replace(/[()（）\[\]{}]/g,'').toLowerCase(); }
function text(v){ return String(v ?? '').replace(/\s+/g,' ').trim(); }
function toNum(v){
  if(typeof v === 'number' && isFinite(v)) return v;
  const s = String(v ?? '').replace(/[,원천\s]/g,'').replace(/[^0-9.\-]/g,'');
  if(!s || s === '-' || s === '.') return 0;
  const n = Number(s); return isFinite(n) ? n : 0;
}
function fmt(n){ if(n === null || n === undefined || n === '') return '-'; return Number(n).toLocaleString('ko-KR') + '원'; }
function table(headers, rows, opts={}){
  if(!rows || !rows.length) return '<p class="muted">표시할 데이터가 없습니다.</p>';
  const head = '<tr>'+headers.map(h=>`<th>${h}</th>`).join('')+'</tr>';
  const body = rows.map(r => '<tr>'+headers.map(h=>{
    const v = r[h] ?? '';
    const cls = typeof v === 'number' || /금액|인원|행|열|페이지/.test(h) ? ' class="num"' : '';
    return `<td${cls}>${escapeHtml(String(v))}</td>`;
  }).join('')+'</tr>').join('');
  return `<div class="scroll ${opts.small?'small':''}"><table>${head}${body}</table></div>`;
}
function escapeHtml(s){return s.replace(/[&<>"]/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));}
function setStatus(msg, ok=true){ const el=$('status'); el.className='status '+(ok?'ok':'err'); el.textContent=msg; }

$('runBtn').addEventListener('click', async () => {
  try{
    setStatus('파일을 읽는 중입니다...', true);
    const excelFile = $('excelFile').files[0];
    const pdfFile = $('pdfFile').files[0];
    if(!excelFile && !pdfFile) throw new Error('엑셀 또는 PDF 파일을 선택해주세요.');
    const report = { createdAt:new Date().toISOString(), excel:null, pdf:null, precheck:[] };
    if(excelFile) report.excel = await parseExcel(excelFile);
    if(pdfFile) report.pdf = await parsePdf(pdfFile);
    report.precheck = buildPrecheck(report);
    state.report = report;
    render(report);
    $('downloadBtn').disabled = false;
    setStatus('값 추출이 완료되었습니다. 아래 표에서 엑셀/PDF가 제대로 읽혔는지 확인해주세요.', true);
  }catch(e){ console.error(e); setStatus('오류가 발생했습니다: '+e.message, false); }
});
$('downloadBtn').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify(state.report,null,2)], {type:'application/json'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download='budget-parser-debug.json'; a.click(); URL.revokeObjectURL(a.href);
});

async function parseExcel(file){
  const originalBuf = await file.arrayBuffer();

  // 1순위: XLSX 내부 XML 직접 파싱. HCell/한컴 계열 파일에서 SheetJS가 빈 시트로 읽는 경우를 보정합니다.
  let rawBook = null;
  try{
    rawBook = await readXlsxRawWorkbook(originalBuf);
  }catch(e){
    console.warn('Raw XLSX parser failed, falling back to SheetJS:', e);
  }

  if(rawBook){
    const visible = rawBook.sheets.filter(s=>!s.hidden);
    const candidates = visible.map(s=>({ ...s, score: sheetScore(s.name) }))
      .filter(s=>s.score>0).sort((a,b)=>b.score-a.score);
    let salary = null;
    const tried = [];
    for(const c of candidates){
      const aoa = await rawBook.getAoa(c.name);
      const parsed = parseSalarySheet(c.name, aoa);
      tried.push({sheet:c.name, score:c.score, found:!!parsed.ok, message:parsed.message || ''});
      if(parsed.ok){ salary = parsed; break; }
    }
    const retireSheets = visible
      .filter(s=>RETIRE_RE.test(s.name))
      .map(async s=>parseRetireSheet(s.name, await rawBook.getAoa(s.name)));
    const retireSheetsResolved = await Promise.all(retireSheets);
    return {
      fileName:file.name,
      parser:'raw-xlsx-xml',
      visibleSheets:visible.map(s=>s.name),
      salaryCandidates:tried,
      salary,
      retirement:retireSheetsResolved
    };
  }

  // 2순위: 일반 XLSX 파일은 SheetJS 경로로 읽습니다.
  const normalizedBuf = await normalizeXlsxForSheetJS(originalBuf);
  const wb = XLSX.read(normalizedBuf, { type:'array', cellDates:false, cellNF:false, cellText:true, raw:false, WTF:false });
  const sheetMeta = (wb.Workbook && wb.Workbook.Sheets) || [];
  const visible = wb.SheetNames.map((name,i)=>({name, hidden: sheetMeta[i]?.Hidden || 0})).filter(s=>!s.hidden);
  const candidates = visible.map(s=>({ ...s, score: sheetScore(s.name) })).filter(s=>s.score>0).sort((a,b)=>b.score-a.score);
  let salary = null;
  const tried = [];
  for(const c of candidates){
    const ws = wb.Sheets[c.name];
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''});
    const parsed = parseSalarySheet(c.name, aoa);
    tried.push({sheet:c.name, score:c.score, found:!!parsed.ok, message:parsed.message || ''});
    if(parsed.ok){ salary = parsed; break; }
  }
  const retireSheets = visible.filter(s=>RETIRE_RE.test(s.name)).map(s=>parseRetireSheet(s.name, XLSX.utils.sheet_to_json(wb.Sheets[s.name], {header:1, raw:false, defval:''})));
  return { fileName:file.name, parser:'sheetjs', visibleSheets:visible.map(s=>s.name), salaryCandidates:tried, salary, retirement:retireSheets };
}

async function readXlsxRawWorkbook(buf){
  if(typeof JSZip === 'undefined') throw new Error('JSZip이 로드되지 않았습니다.');
  const zip = await JSZip.loadAsync(buf);
  const workbookFile = zip.file('xl/workbook.xml');
  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if(!workbookFile || !relsFile) throw new Error('workbook.xml을 찾지 못했습니다.');
  const parser = new DOMParser();
  const workbookXml = parser.parseFromString(await workbookFile.async('string'), 'application/xml');
  const relsXml = parser.parseFromString(await relsFile.async('string'), 'application/xml');
  const relMap = {};
  localElements(relsXml, 'Relationship').forEach(el=>{
    relMap[el.getAttribute('Id')] = el.getAttribute('Target');
  });

  const sharedStrings = await readSharedStrings(zip, parser);
  const sheets = localElements(workbookXml, 'sheet').map(el=>{
    const rid = el.getAttribute('r:id') || el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    const target = relMap[rid] || '';
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\.\//,'');
    const state = el.getAttribute('state') || '';
    return {
      name: el.getAttribute('name'),
      rid,
      path,
      hidden: state === 'hidden' || state === 'veryHidden'
    };
  }).filter(s=>s.name && s.path);

  const cache = new Map();
  const getAoa = async (sheetName) => {
    if(cache.has(sheetName)) return cache.get(sheetName);
    const meta = sheets.find(s=>s.name===sheetName);
    if(!meta) return [];
    const aoa = await sheetXmlToAoa(zip, meta.path, sharedStrings, parser);
    cache.set(sheetName, aoa);
    return aoa;
  };

  return {sheets, getAoa};
}

async function readSharedStrings(zip, parser){
  const f = zip.file('xl/sharedStrings.xml');
  if(!f) return [];
  const xml = parser.parseFromString(await f.async('string'), 'application/xml');
  return localElements(xml, 'si').map(si => localElements(si, 't').map(t=>t.textContent || '').join(''));
}

function localElements(root, localName){
  return Array.from(root.getElementsByTagName('*')).filter(el=>el.localName === localName);
}

async function sheetXmlToAoa(zip, path, sharedStrings, parser){
  const f = zip.file(path);
  if(!f) return [];
  const xml = parser.parseFromString(await f.async('string'), 'application/xml');
  const rows = localElements(xml, 'row');
  const aoa = [];
  for(const row of rows){
    const rIdx = Number(row.getAttribute('r') || (aoa.length+1)) - 1;
    if(!aoa[rIdx]) aoa[rIdx] = [];
    const cells = Array.from(row.children).filter(el=>el.localName === 'c');
    for(const c of cells){
      const ref = c.getAttribute('r') || '';
      const cIdx = ref ? colRefToIndex(ref) : aoa[rIdx].length;
      aoa[rIdx][cIdx] = cellValue(c, sharedStrings);
    }
  }
  return aoa.map(r => r || []);
}
function colRefToIndex(ref){
  const letters = String(ref).match(/[A-Z]+/i)?.[0]?.toUpperCase() || 'A';
  let n = 0;
  for(const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function cellValue(cell, sharedStrings){
  const t = cell.getAttribute('t');
  const vEl = Array.from(cell.children).find(el=>el.localName === 'v');
  const fEl = Array.from(cell.children).find(el=>el.localName === 'f');
  if(t === 'inlineStr'){
    return localElements(cell, 't').map(x=>x.textContent || '').join('');
  }
  const raw = vEl ? (vEl.textContent || '') : '';
  if(t === 's') return sharedStrings[Number(raw)] ?? '';
  if(t === 'str') return raw;
  if(raw !== ''){
    const n = Number(raw);
    return Number.isFinite(n) && /^-?\d+(\.\d+)?$/.test(raw) ? n : raw;
  }
  // 수식 셀에 계산 캐시가 없으면 수식 문자열을 남겨 디버그에서 확인할 수 있게 합니다.
  return fEl ? '=' + (fEl.textContent || '') : '';
}

async function normalizeXlsxForSheetJS(buf){
  // 일부 예산서 파일은 엑셀 XML 태그가 <x:worksheet>, <x:c>처럼 접두어가 붙어 저장됩니다.
  // SheetJS 0.18.x가 이 형식을 빈 시트로 읽는 경우가 있어, 브라우저 안에서만 표준 형태로 정규화합니다.
  if(typeof JSZip === 'undefined') return buf;
  try{
    const zip = await JSZip.loadAsync(buf);
    const xmlFiles = Object.keys(zip.files).filter(p => /(^xl\/.*\.xml$)|(^xl\/_rels\/.*\.rels$)|(^_rels\/.*\.rels$)/.test(p));
    let changed = false;
    for(const path of xmlFiles){
      const file = zip.file(path);
      if(!file) continue;
      let xml = await file.async('string');
      const before = xml;
      // x: 접두어만 제거합니다. r:, mc:, a: 등 관계/드로잉 접두어는 유지합니다.
      xml = xml.replace(/<\/?x:/g, m => m.replace('x:', ''));
      xml = xml.replace(/\sxmlns:x="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/g,
                        ' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
      if(xml !== before){ zip.file(path, xml); changed = true; }
    }
    if(!changed) return buf;
    return await zip.generateAsync({type:'arraybuffer', compression:'DEFLATE'});
  }catch(e){
    console.warn('XLSX namespace normalization skipped:', e);
    return buf;
  }
}

function sheetScore(name){
  let score = 0; const n = norm(name);
  if(/교직원보수일람표/.test(n)) score += 100;
  if(/년|연/.test(n)) score += 30;
  if(/월/.test(n)) score -= 20;
  if(SALARY_SHEET_RE.test(name)) score += 40;
  if(/월급여|비월정|간이세액|세출|세입/.test(name)) score -= 100;
  return score;
}
function parseSalarySheet(sheetName, aoa){
  const preview = aoa.slice(0,40).map((r,i)=>({행:i+1, 내용:r.map(text).filter(Boolean).join(' | ').slice(0,250)})).filter(x=>x.내용);
  if(!aoa.length) return {ok:false, sheetName, message:'빈 시트', preview};
  const maxRows = Math.min(80, aoa.length), maxCols = Math.max(...aoa.slice(0,maxRows).map(r=>r.length),0);
  const colText = Array.from({length:maxCols},(_,c)=>aoa.slice(0,maxRows).map(r=>r[c]).map(text).join(' '));
  const jobCol = findCol(colText, [/^직명$/, /직명/]);
  const nameCol = findCol(colText, [/^성명$/, /성명/]);
  const baseCol = findCol(colText, [/본봉/, /기본급/]);
  const totalCol = findCol(colText, [/지급액계/, /지급액.*계/, /월지급액계/, /합계/]);
  if(jobCol < 0 || baseCol < 0 || totalCol < 0) return {ok:false, sheetName, message:`헤더 미발견: 직명=${jobCol+1}, 본봉=${baseCol+1}, 지급액계=${totalCol+1}`, preview};
  const headerRow = findHeaderRow(aoa, [jobCol, baseCol, totalCol]);
  const allowanceCols = [];
  for(let c=baseCol+1; c<totalCol; c++){
    const h = headerNameForCol(aoa, c, headerRow) || cleanHeader(colText[c]);
    if(h && !/성명|호봉|직명|일련|본봉|기본급|지급액계|합계|소계/.test(h)) allowanceCols.push({ col:c, name:h || `열${c+1}` });
  }
  const teacherEnd = findRow(aoa, headerRow+1, /소계\s*\(?교원\)?|소계교원/);
  const staffEnd = findRow(aoa, (teacherEnd>=0?teacherEnd+1:headerRow+1), /소계\s*\(?(직원|일반직)\)?|소계직원|소계일반직/);
  const teacherRows = rowsInRange(aoa, headerRow+1, teacherEnd>=0?teacherEnd:staffEnd, {jobCol,nameCol,baseCol,totalCol,allowanceCols});
  const staffRows = rowsInRange(aoa, teacherEnd>=0?teacherEnd+1:headerRow+1, staffEnd, {jobCol,nameCol,baseCol,totalCol,allowanceCols});
  const teacherSubtotal = teacherEnd>=0 ? parseDataRow(aoa[teacherEnd], {jobCol,nameCol,baseCol,totalCol,allowanceCols}) : null;
  const staffSubtotal = staffEnd>=0 ? parseDataRow(aoa[staffEnd], {jobCol,nameCol,baseCol,totalCol,allowanceCols}) : null;
  const summary = summarizeSalary(teacherRows, staffRows, teacherSubtotal, staffSubtotal, allowanceCols);
  return {ok:true, sheetName, header:{headerRow:headerRow+1, jobCol:jobCol+1, nameCol:nameCol+1, baseCol:baseCol+1, totalCol:totalCol+1, allowanceCols:allowanceCols.map(x=>({열:x.col+1, 이름:x.name}))}, ranges:{teacherStart:headerRow+2, teacherEnd:teacherEnd>=0?teacherEnd: null, staffStart:teacherEnd>=0?teacherEnd+2:headerRow+2, staffEnd:staffEnd>=0?staffEnd:null}, summary, teacherRows, staffRows, teacherSubtotal, staffSubtotal, preview};
}
function findCol(colText, patterns){
  let best=-1, bestScore=-1;
  colText.forEach((t,i)=>{ const n=norm(t); patterns.forEach((re,idx)=>{ if(re.test(n) || re.test(t)){ const sc=100-idx; if(sc>bestScore){best=i; bestScore=sc;} } }); });
  return best;
}
function findHeaderRow(aoa, cols){
  let best=0, score=-1;
  aoa.slice(0,80).forEach((r,i)=>{ let s=0; cols.forEach(c=>{ if(text(r[c])) s++; }); const all=r.map(text).join(' '); if(/직명/.test(all))s+=2; if(/본봉|기본급/.test(all))s+=2; if(/지급액계/.test(all))s+=2; if(s>score){score=s;best=i;} }); return best;
}
function findRow(aoa, start, re){ for(let r=Math.max(0,start); r<aoa.length; r++){ if(re.test(aoa[r].map(text).join(' '))) return r; } return -1; }
function rowsInRange(aoa, start, end, cfg){
  const last = end>=0 ? end : aoa.length;
  const rows=[];
  for(let r=Math.max(0,start); r<last; r++){
    const row = parseDataRow(aoa[r], cfg); if(row && (row.직명 || row.성명 || row.본봉 || row.지급액계)) rows.push({...row, 행:r+1});
  }
  return rows.filter(r=> !/소계|합계|계\s*$/.test(`${r.직명} ${r.성명}`));
}
function parseDataRow(row, cfg){
  if(!row) return null;
  const out = {직명:text(row[cfg.jobCol]), 성명:cfg.nameCol>=0?text(row[cfg.nameCol]):'', 본봉:toNum(row[cfg.baseCol]), 지급액계:toNum(row[cfg.totalCol])};
  out.수당 = {};
  cfg.allowanceCols.forEach(a=>{ out.수당[a.name]=toNum(row[a.col]); });
  return out;
}
function summarizeSalary(teacherRows, staffRows, teacherSubtotal, staffSubtotal, allowanceCols){
  const sumRows = (rows) => ({인원:rows.filter(r=>r.본봉 || r.지급액계 || Object.values(r.수당).some(Boolean)).length, 본봉:rows.reduce((s,r)=>s+r.본봉,0), 지급액계:rows.reduce((s,r)=>s+r.지급액계,0)});
  const t = sumRows(teacherRows), s = sumRows(staffRows);
  const allowances = allowanceCols.map(a=>({항목:a.name, 교원금액:teacherRows.reduce((sum,r)=>sum+(r.수당[a.name]||0),0), 직원금액:staffRows.reduce((sum,r)=>sum+(r.수당[a.name]||0),0)})).filter(x=>x.교원금액||x.직원금액);
  return {교원:{...t, 소계본봉:teacherSubtotal?.본봉||0, 소계지급액계:teacherSubtotal?.지급액계||0}, 직원:{...s, 소계본봉:staffSubtotal?.본봉||0, 소계지급액계:staffSubtotal?.지급액계||0}, 수당:allowances};
}
function cleanHeader(s){ return text(s).replace(/\([A-Z]\)/g,'').replace(/\s+/g,' ').trim().slice(0,40); }
function headerNameForCol(aoa, col, headerRow){
  const start = Math.max(0, headerRow - 6);
  const end = Math.min(aoa.length - 1, headerRow + 2);
  const vals = [];
  for(let r=start; r<=end; r++){
    const v = text(aoa[r]?.[col]);
    if(!v) continue;
    if(toNum(v) || /소계|합계|원장|교사|직원|성명|직명|호봉|순번|번호|구분/.test(v)) continue;
    vals.push(v);
  }
  return cleanHeader([...new Set(vals)].join(' '));
}
function parseRetireSheet(sheetName, aoa){
  let positive = 0, cells=[];
  aoa.forEach((r,ri)=>r.forEach((v,ci)=>{ const n=toNum(v); const around=r.map(text).join(' '); if(n>0 && !/연도|년월|날짜/.test(around)){ positive += n; cells.push({행:ri+1, 열:ci+1, 값:n, 주변:around.slice(0,160)}); } }));
  return {sheetName, positiveAmount:positive, hasRetirementAmount:positive>0, positiveCells:cells.slice(0,30)};
}

async function parsePdf(file){
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({data:buf}).promise;
  const pages = [], lines=[];
  for(let p=1;p<=pdf.numPages;p++){
    const page = await pdf.getPage(p); const content = await page.getTextContent();
    const items = content.items.map(it=>({str:it.str, x:Math.round(it.transform[4]), y:Math.round(it.transform[5])})).filter(it=>text(it.str));
    const grouped = groupPdfLines(items).map(l=>({...l, page:p}));
    pages.push({page:p, lineCount:grouped.length}); lines.push(...grouped);
  }
  const items = extractBudgetItems(lines);
  return { fileName:file.name, pageCount:pdf.numPages, lines:lines.map(l=>({페이지:l.page, y:l.y, 텍스트:l.text})), items };
}
function groupPdfLines(items){
  const buckets = new Map();
  items.forEach(it=>{ const key=Math.round(it.y/3)*3; if(!buckets.has(key)) buckets.set(key,[]); buckets.get(key).push(it); });
  return [...buckets.entries()].sort((a,b)=>b[0]-a[0]).map(([y,arr])=>({y, text:arr.sort((a,b)=>a.x-b.x).map(x=>x.str).join(' ').replace(/\s+/g,' ').trim()})).filter(l=>l.text);
}
function extractBudgetItems(lines){
  const out=[];
  let currentTotal = null;
  for(let i=0;i<lines.length;i++){
    const l=lines[i], t=l.text;
    const totalRow = parseTotalBudgetRow(t);
    if(totalRow){
      currentTotal = {항목:totalRow.항목, 페이지:l.page, 행:i+1, 금액:totalRow.금액};
      out.push({페이지:l.page, 행:i+1, 상위항목:'', 항목:totalRow.항목, PDF금액:totalRow.금액, PDF금액천원:Math.round(totalRow.금액/1000), 인원:null, 산출기초:'', 구분:'총액행'});
      continue;
    }
    const calc = parseCalcLine(t);
    if(calc){
      const name = findCalcName(lines, i) || currentTotal?.항목 || '산출항목미상';
      out.push({페이지:l.page, 행:i+1, 상위항목:currentTotal?.항목 || '', 항목:name, PDF금액:calc.amount, PDF금액천원:Math.round(calc.amount/1000), 인원:calc.people, 산출기초:t, 구분:'산출기초'});
    }
  }
  return out.filter(x=>!LEGAL_RE.test(x.항목));
}
function parseTotalBudgetRow(t){
  if(/산출|예산구분|발행일|과\s*목|보조금|수익자/.test(t)) return null;
  const m = t.match(/^\s*([^0-9=]{2,40}?)\s+([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\s+([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\s+([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\s+([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\s*$/);
  if(!m) return null;
  const name = text(m[1]).replace(/\s+/g,'').trim();
  if(!name || /본예산|단위|천원/.test(name)) return null;
  return {항목:name, 금액:toNum(m[5])*1000};
}
function parseCalcLine(t){
  const compact = t.replace(/\s+/g,'');
  let m = compact.match(/([0-9,]+)원\*([0-9,]+)명\*([0-9,]+)월=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:toNum(m[2]), months:toNum(m[3]), amount:toNum(m[4])};
  m = compact.match(/([0-9,]+)원\*([0-9,]+)월=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:null, months:toNum(m[2]), amount:toNum(m[3])};
  m = compact.match(/([0-9,]+)원\*([0-9,]+)명\*1월=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:toNum(m[2]), months:1, amount:toNum(m[3])};
  return null;
}
function findCalcName(lines, idx){
  for(let j=idx-1; j>=Math.max(0, idx-6); j--){
    let s = text(lines[j]?.text || '');
    s = s.replace(/\(본예산\)|\(보조금및지원금\)|\(수익자부담금\)|\(그밖의수입\)/g,' ')
         .replace(/교\s*$/,'').replace(/환\s*$/,'').replace(/운\s*$/,'').replace(/경\s*$/,'')
         .replace(/\s+/g,' ').trim();
    if(!s) continue;
    if(/[0-9,]+원\s*\*/.test(s)) continue;
    if(/예산구분|발행일|산출|과\s*목/.test(s)) continue;
    const cleaned = s.replace(/\s+/g,'');
    if(cleaned.length >= 2 && cleaned.length <= 40) return cleaned;
  }
  return '';
}

function buildPrecheck(report){
  const rows=[];
  const issues=[];
  const ex = report.excel?.salary?.summary;
  const pdfItems = report.pdf?.items || [];
  const totals = pdfItems.filter(x=>x.구분==='총액행');
  const calcs = pdfItems.filter(x=>x.구분==='산출기초');
  const totalByName = (name) => totals.find(x=>norm(x.항목)===norm(name))?.PDF금액 || 0;
  const calcByName = (re) => calcs.filter(x=>re.test(x.항목));
  const amountByCalc = (re) => calcByName(re).reduce((s,x)=>s+x.PDF금액,0);
  const peopleByCalc = (re) => calcByName(re).reduce((s,x)=>s+(x.인원||0),0);
  const verdict = (excelAmount, pdfAmount, excelPeople, pdfPeople) => {
    const amountOk = Number(excelAmount||0) === Number(pdfAmount||0);
    const peopleOk = pdfPeople === '' || pdfPeople === null || pdfPeople === undefined ? true : Number(excelPeople||0) === Number(pdfPeople||0);
    if(amountOk && peopleOk) return '일치';
    const parts=[]; if(!amountOk) parts.push('금액 차이'); if(!peopleOk) parts.push('인원 차이'); return parts.join(', ');
  };
  if(ex){
    const teacherBase = ex.교원.소계본봉 || ex.교원.본봉;
    const staffBase = ex.직원.소계본봉 || ex.직원.본봉;
    const teacherPayTotal = totalByName('교원급여') || amountByCalc(/^교원급여$/);
    const teacherPayPeople = peopleByCalc(/^교원급여$/) || '';
    rows.push({구분:'급여', 항목:'교원급여', 엑셀금액:teacherBase, 엑셀인원:ex.교원.인원, PDF금액:teacherPayTotal, PDF인원:teacherPayPeople, 확인:verdict(teacherBase, teacherPayTotal, ex.교원.인원, teacherPayPeople)});

    const afterTeacherPay = amountByCalc(/방과후.*교원.*급여|방과후교원급여/);
    const afterTeacherPeople = peopleByCalc(/방과후.*교원.*급여|방과후교원급여/) || '';
    if(afterTeacherPay) rows.push({구분:'급여', 항목:'방과후교원급여', 엑셀금액:'', 엑셀인원:'', PDF금액:afterTeacherPay, PDF인원:afterTeacherPeople, 확인:'PDF 산출기초 확인'});

    const staffPayTotal = totalByName('직원급여') || amountByCalc(/사무직원급여|조리직원급여|보조교사급여|차량기사급여|차량보조급여|환경미화원급여/);
    const staffPayPeople = peopleByCalc(/사무직원급여|조리직원급여|보조교사급여|차량기사급여|차량보조급여|환경미화원급여/) || '';
    rows.push({구분:'급여', 항목:'직원급여', 엑셀금액:staffBase, 엑셀인원:ex.직원.인원, PDF금액:staffPayTotal, PDF인원:staffPayPeople, 확인:verdict(staffBase, staffPayTotal, ex.직원.인원, '')});

    const teacherAllowanceTotal = totalByName('교원수당');
    const staffAllowanceTotal = totalByName('직원수당');
    const allowanceRows = ex.수당.filter(a=>!LEGAL_RE.test(a.항목));
    const teacherAllowanceExcel = allowanceRows.reduce((s,a)=>s+a.교원금액,0);
    const staffAllowanceExcel = allowanceRows.reduce((s,a)=>s+a.직원금액,0);
    rows.push({구분:'수당', 항목:'교원수당', 엑셀금액:teacherAllowanceExcel, 엑셀인원:'', PDF금액:teacherAllowanceTotal, PDF인원:'', 확인:allowanceVerdict('교원', allowanceRows, pdfItems, teacherAllowanceTotal)});
    rows.push({구분:'수당', 항목:'직원수당', 엑셀금액:staffAllowanceExcel, 엑셀인원:'', PDF금액:staffAllowanceTotal, PDF인원:'', 확인:allowanceVerdict('직원', allowanceRows, pdfItems, staffAllowanceTotal)});

    const car = allowanceRows.find(a=>/자가운전|자가.*운전|차량보조|차량.*보조/.test(a.항목));
    if(car && (car.교원금액 || car.직원금액)){
      const pdfHasCar = pdfItems.some(x=>/자가운전|자가.*운전/.test(x.항목));
      if(!pdfHasCar) issues.push({번호:issues.length+1, 지적내용:'교원 및 직원 자가운전보조금 예산 미편성', 근거:`엑셀 자가운전보조금: 교원 ${fmt(car.교원금액)}, 직원 ${fmt(car.직원금액)} / PDF 개별 항목 없음`});
    }
  }

  const vehicleWrong = calcs.find(x=>/차량기사급여/.test(x.항목) && /통학차량이용비|차량이용비|차량/.test(x.상위항목));
  if(vehicleWrong){
    issues.push({번호:issues.length+1, 지적내용:`차량기사급여 ${Math.round(vehicleWrong.PDF금액/10000).toLocaleString('ko-KR')}만원을 직원인건비가 아닌 ${vehicleWrong.상위항목}의 산출내역에 편성`, 근거:`PDF ${vehicleWrong.페이지}쪽: ${vehicleWrong.산출기초}`});
  }

  const excelRetire = (report.excel?.retirement || []).some(r=>r.hasRetirementAmount);
  const pdfRetire = pdfItems.some(x=>/퇴직.*적립|퇴직금|퇴직급여|퇴직충당/.test(x.항목));
  if(excelRetire && !pdfRetire){
    issues.push({번호:issues.length+1, 지적내용:'엑셀에는 퇴직 적립금액이 있으나 퇴직적립금 미편성', 근거:'퇴직 관련 엑셀 시트에는 0원을 초과하는 금액이 있으나 PDF에서 퇴직 관련 편성 항목을 찾지 못했습니다.'});
  }
  return {rows, issues};
}
function allowanceVerdict(kind, allowanceRows, pdfItems, groupTotal){
  const amountField = kind === '교원' ? '교원금액' : '직원금액';
  const relevant = allowanceRows.filter(a=>a[amountField]);
  const direct = [];
  const missing = [];
  for(const a of relevant){
    const key = allowanceKey(a.항목);
    const hit = key && pdfItems.some(x=>allowanceKey(x.항목).includes(key) || key.includes(allowanceKey(x.항목)));
    (hit ? direct : missing).push(a);
  }
  const missingSum = missing.reduce((s,a)=>s+a[amountField],0);
  if(!groupTotal && missingSum) return `추가확인필요: PDF ${kind}수당 총액행 없음`;
  if(missingSum && missingSum === groupTotal) return `통합편성(${missing.map(a=>`${a.항목} ${fmt(a[amountField])}`).join(' + ')} = ${fmt(missingSum)})`;
  if(relevant.reduce((s,a)=>s+a[amountField],0) === groupTotal) return '총액 일치';
  if(missingSum) return `추가확인필요: 미개별표시 수당 합계 ${fmt(missingSum)}, PDF ${kind}수당 ${fmt(groupTotal)}`;
  return '개별 또는 총액 확인';
}
function allowanceKey(s){
  return norm(s).replace(/수당|보조금|지원비|비/g,'');
}

function render(report){
  if(report.excel){
    const ex=report.excel;
    $('excelSummary').innerHTML = `<span class="pill">파일 ${escapeHtml(ex.fileName)}</span><span class="pill">보이는 시트 ${ex.visibleSheets.length}개</span>`;
    let html = '<h3 class="section-title">시트 후보</h3>'+table(['sheet','score','found','message'], ex.salaryCandidates, {small:true});
    if(ex.salary?.ok){
      html += `<h3 class="section-title">선택된 보수 시트: ${escapeHtml(ex.salary.sheetName)}</h3>`;
      html += table(['항목','값'], Object.entries(ex.salary.header).map(([항목,값])=>({항목,값:JSON.stringify(값)})), {small:true});
      html += '<h3 class="section-title">보수 요약</h3>'+table(['구분','본봉','소계본봉','지급액계','소계지급액계','인원'], [
        {구분:'교원', 본봉:fmt(ex.salary.summary.교원.본봉), 소계본봉:fmt(ex.salary.summary.교원.소계본봉), 지급액계:fmt(ex.salary.summary.교원.지급액계), 소계지급액계:fmt(ex.salary.summary.교원.소계지급액계), 인원:ex.salary.summary.교원.인원},
        {구분:'직원', 본봉:fmt(ex.salary.summary.직원.본봉), 소계본봉:fmt(ex.salary.summary.직원.소계본봉), 지급액계:fmt(ex.salary.summary.직원.지급액계), 소계지급액계:fmt(ex.salary.summary.직원.소계지급액계), 인원:ex.salary.summary.직원.인원}
      ]);
      html += '<h3 class="section-title">수당 추출</h3>'+table(['항목','교원금액','직원금액'], ex.salary.summary.수당.map(a=>({항목:a.항목, 교원금액:fmt(a.교원금액), 직원금액:fmt(a.직원금액)})));
      html += '<h3 class="section-title">교원 행 미리보기</h3>'+table(['행','직명','성명','본봉','지급액계'], ex.salary.teacherRows.slice(0,40).map(r=>({...r,본봉:fmt(r.본봉),지급액계:fmt(r.지급액계)})), {small:true});
      html += '<h3 class="section-title">직원 행 미리보기</h3>'+table(['행','직명','성명','본봉','지급액계'], ex.salary.staffRows.slice(0,40).map(r=>({...r,본봉:fmt(r.본봉),지급액계:fmt(r.지급액계)})), {small:true});
    } else {
      html += '<p class="bad">보수 시트를 구조적으로 읽지 못했습니다. 후보 시트의 상단 미리보기를 확인해주세요.</p>';
      const first = ex.salaryCandidates[0]?.sheet;
      if(first){ /* preview unavailable here if failed for nonselected; use candidate message only */ }
    }
    html += '<h3 class="section-title">퇴직 관련 시트</h3>'+table(['sheetName','positiveAmount','hasRetirementAmount'], ex.retirement.map(r=>({sheetName:r.sheetName, positiveAmount:fmt(r.positiveAmount), hasRetirementAmount:r.hasRetirementAmount?'있음':'없음'})));
    $('excelTables').innerHTML = html;
  }
  if(report.pdf){
    const pdf=report.pdf;
    $('pdfSummary').innerHTML = `<span class="pill">파일 ${escapeHtml(pdf.fileName)}</span><span class="pill">${pdf.pageCount}페이지</span><span class="pill">항목 ${pdf.items.length}개</span>`;
    let html = '<h3 class="section-title">PDF 추출 항목</h3>'+table(['구분','페이지','항목','PDF금액','PDF금액천원','인원','산출기초'], pdf.items.map(x=>({...x,PDF금액:fmt(x.PDF금액)})));
    html += '<h3 class="section-title">PDF 원문 라인</h3>'+table(['페이지','y','텍스트'], pdf.lines.slice(0,300), {small:true});
    $('pdfTables').innerHTML = html;
  }
  const review = report.precheck || {rows:[], issues:[]};
  let reviewHtml = '<h3 class="section-title">금액 및 인원 검토</h3>' + table(['구분','항목','엑셀금액','엑셀인원','PDF금액','PDF인원','확인'], review.rows.map(r=>({...r,엑셀금액:fmt(r.엑셀금액),PDF금액:fmt(r.PDF금액)})));
  reviewHtml += '<h3 class="section-title">지적사항</h3>' + table(['번호','지적내용','근거'], review.issues || []);
  $('precheck').innerHTML = reviewHtml;
}
