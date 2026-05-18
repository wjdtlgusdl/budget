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
function closeMoney(a,b,tol=1000){
  a = Number(a||0); b = Number(b||0);
  if(!a && !b) return true;
  return Math.abs(a-b) <= tol;
}
function amountDiff(a,b){ return Math.round(Number(a||0) - Number(b||0)); }
function shortWon(n){ return Math.abs(n) >= 10000 ? Math.round(n/10000).toLocaleString('ko-KR') + '만원' : fmt(n); }
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

  // v34 핵심 변경:
  // - 한컴/HCell 계열 XLSX는 SheetJS가 빈 시트처럼 읽는 경우가 있어 raw-xlsx-xml 파서가 필요합니다.
  // - 반대로 MS Excel/기관별 서식 파일은 SheetJS가 더 안정적으로 읽는 경우가 있습니다.
  // - 따라서 두 경로를 모두 시도하고, 보수표 구조를 실제로 읽은 쪽을 선택합니다.
  const attempts = [];

  // A. SheetJS 경로 먼저 시도: 일반 Excel 파일, 넓은 열 범위(IX 등), 표준 sharedStrings 파일에 강함
  try{
    const sheetReport = await parseExcelWithSheetJS(file.name, originalBuf);
    attempts.push(sheetReport);
  }catch(e){
    attempts.push({parser:'sheetjs', error:e.message, salaryCandidates:[], salary:null, visibleSheets:[]});
    console.warn('SheetJS parse failed:', e);
  }

  // B. raw XML 경로 시도: HCell/한컴 계열에서 SheetJS가 빈 시트로 읽는 문제 보정
  try{
    const rawReport = await parseExcelWithRawXml(file.name, originalBuf);
    attempts.push(rawReport);
  }catch(e){
    attempts.push({parser:'raw-xlsx-xml', error:e.message, salaryCandidates:[], salary:null, visibleSheets:[]});
    console.warn('Raw XLSX parser failed:', e);
  }

  // C. 선택 기준: salary가 정상인 결과 우선, 그 다음 후보 진단이 더 풍부한 결과
  const successful = attempts.filter(x=>x && x.salary && x.salary.ok);
  let chosen = successful[0];
  if(successful.length > 1){
    // (년) 시트/보수일람표를 읽은 결과를 우선
    chosen = successful.sort((a,b)=>{
      const as = selectedSalaryScore(a), bs = selectedSalaryScore(b);
      return bs-as;
    })[0];
  }
  if(!chosen){
    chosen = attempts.sort((a,b)=>(b.salaryCandidates?.length||0)-(a.salaryCandidates?.length||0))[0] || {fileName:file.name};
  }
  chosen.fileName = file.name;
  chosen.parserAttempts = attempts.map(a=>({
    parser:a.parser,
    error:a.error||'',
    visibleSheets:a.visibleSheets||[],
    salaryFound:!!(a.salary&&a.salary.ok),
    salarySheet:a.salary?.sheetName||'',
    candidates:a.salaryCandidates||[]
  }));
  return chosen;
}

function selectedSalaryScore(report){
  const name = report.salary?.sheetName || '';
  let score = 0;
  if(/교직원보수일람표/.test(name)) score += 100;
  if(/\(년\)|년간|연간/.test(name)) score += 50;
  if(/\(월\)|월간|월급여/.test(name)) score -= 20;
  if(report.parser === 'raw-xlsx-xml') score += 5;
  if(report.parser === 'sheetjs') score += 3;
  return score;
}

async function parseExcelWithRawXml(fileName, originalBuf){
  const rawBook = await readXlsxRawWorkbook(originalBuf);
  const visible = rawBook.sheets.filter(s=>!s.hidden);
  const candidates = visible.map(s=>({ ...s, score: sheetScore(s.name) }))
    .filter(s=>s.score>0).sort((a,b)=>b.score-a.score);
  let salary = null;
  const tried = [];
  for(const c of candidates){
    const aoa = await rawBook.getAoa(c.name);
    const parsed = parseSalarySheet(c.name, aoa);
    tried.push({sheet:c.name, score:c.score, found:!!parsed.ok, message:parsed.message || '', parser:'raw-xlsx-xml'});
    if(parsed.ok){ salary = parsed; break; }
  }
  const retireSheets = visible
    .filter(s=>RETIRE_RE.test(s.name))
    .map(async s=>parseRetireSheet(s.name, await rawBook.getAoa(s.name)));
  const retireSheetsResolved = await Promise.all(retireSheets);
  return {
    fileName,
    parser:'raw-xlsx-xml',
    visibleSheets:visible.map(s=>s.name),
    salaryCandidates:tried,
    salary,
    retirement:retireSheetsResolved
  };
}

async function parseExcelWithSheetJS(fileName, originalBuf){
  const normalizedBuf = await normalizeXlsxForSheetJS(originalBuf);
  const wb = XLSX.read(normalizedBuf, {
    type:'array',
    cellDates:false,
    cellNF:false,
    cellText:true,
    raw:false,
    WTF:false,
    dense:false
  });
  const sheetMeta = (wb.Workbook && wb.Workbook.Sheets) || [];
  const visible = wb.SheetNames.map((name,i)=>({name, hidden: sheetMeta[i]?.Hidden || 0})).filter(s=>!s.hidden);
  const candidates = visible.map(s=>({ ...s, score: sheetScore(s.name) })).filter(s=>s.score>0).sort((a,b)=>b.score-a.score);
  let salary = null;
  const tried = [];
  for(const c of candidates){
    const ws = wb.Sheets[c.name];
    const aoa = sheetToAoaRobust(ws);
    const parsed = parseSalarySheet(c.name, aoa);
    tried.push({sheet:c.name, score:c.score, found:!!parsed.ok, message:parsed.message || '', parser:'sheetjs'});
    if(parsed.ok){ salary = parsed; break; }
  }
  const retireSheets = visible
    .filter(s=>RETIRE_RE.test(s.name))
    .map(s=>parseRetireSheet(s.name, sheetToAoaRobust(wb.Sheets[s.name])));
  return { fileName, parser:'sheetjs', visibleSheets:visible.map(s=>s.name), salaryCandidates:tried, salary, retirement:retireSheets };
}

function sheetToAoaRobust(ws){
  if(!ws || !ws['!ref']) return [];
  // defval을 유지하되, 매우 넓은 IX 범위 같은 시트에서도 필요한 실제 값은 빠지지 않도록 SheetJS 기본 변환 사용
  let aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''});
  // 어떤 셀은 .w에는 표시값이 있고 .v가 빈 경우가 있어 직접 보완합니다.
  const range = XLSX.utils.decode_range(ws['!ref']);
  for(let R=range.s.r; R<=Math.min(range.e.r, 120); ++R){
    if(!aoa[R]) aoa[R] = [];
    for(let C=range.s.c; C<=Math.min(range.e.c, 260); ++C){
      if(aoa[R][C] !== undefined && aoa[R][C] !== '') continue;
      const addr = XLSX.utils.encode_cell({r:R,c:C});
      const cell = ws[addr];
      if(cell){
        aoa[R][C] = cell.w ?? cell.v ?? '';
      }
    }
  }
  return aoa;
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
  const raw = await f.async('string');

  // v36: 한컴/HCell 파일은 sharedStrings가 x: 접두어로 저장되는 경우가 많습니다.
  // DOMParser 결과가 비어 있거나 일부만 읽히면 헤더가 전부 숫자 인덱스로 남아
  // '직명/본봉/지급액계 헤더 미발견'이 됩니다.
  // 그래서 정규식 파서를 항상 먼저 만들어 두고, DOM 결과보다 풍부하면 정규식 결과를 사용합니다.
  const regexResult = [];
  const siRe = /<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g;
  let m;
  while((m = siRe.exec(raw))){
    const block = m[1];
    const parts = [];
    const tRe = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
    let tm;
    while((tm = tRe.exec(block))) parts.push(xmlDecode(tm[1]));
    regexResult.push(parts.join(''));
  }

  let domResult = [];
  try{
    const xml = parser.parseFromString(raw, 'application/xml');
    domResult = localElements(xml, 'si').map(si => localElements(si, 't').map(t=>t.textContent || '').join(''));
  }catch(e){
    console.warn('sharedStrings DOM parse failed, using regex fallback', e);
  }

  const domUseful = domResult.filter(Boolean).length;
  const regexUseful = regexResult.filter(Boolean).length;
  return regexUseful >= domUseful ? regexResult : domResult;
}

function localElements(root, localName){
  return Array.from(root.getElementsByTagName('*')).filter(el=>el.localName === localName);
}

async function sheetXmlToAoa(zip, path, sharedStrings, parser){
  const f = zip.file(path);
  if(!f) return [];
  const raw = await f.async('string');
  let aoa = [];
  try{
    const xml = parser.parseFromString(raw, 'application/xml');
    const rows = localElements(xml, 'row');
    for(const row of rows){
      const rIdx = Number(row.getAttribute('r') || (aoa.length+1)) - 1;
      if(!aoa[rIdx]) aoa[rIdx] = [];
      const cells = localElements(row, 'c');
      for(const c of cells){
        const ref = c.getAttribute('r') || '';
        const cIdx = ref ? colRefToIndex(ref) : aoa[rIdx].length;
        aoa[rIdx][cIdx] = cellValue(c, sharedStrings);
      }
    }
    aoa = aoa.map(r => r || []);
  }catch(e){
    console.warn('sheet XML DOM parse failed, using regex fallback', e);
  }
  // DOMParser가 빈 값처럼 읽었거나 shared string이 해석되지 않은 경우, 원문 XML 기반 파서로 재시도합니다.
  // v36: 헤더 점수뿐 아니라 실제 텍스트 개수도 비교합니다. 일부 파일은 DOM AOA가 숫자 인덱스만
  // 채워져 점수는 0이지만 행 수는 존재하므로, 정규식 AOA가 더 많은 텍스트를 복원하면 그것을 채택합니다.
  const regexAoa = sheetXmlToAoaByRegex(raw, sharedStrings);
  const domScore = aoaScoreForHeaders(aoa);
  const regexScore = aoaScoreForHeaders(regexAoa);
  const domText = countUsefulTextCells(aoa);
  const regexText = countUsefulTextCells(regexAoa);
  if(regexScore > domScore || (regexScore === domScore && regexText > domText)){
    aoa = regexAoa;
  }
  return aoa;
}

function countUsefulTextCells(aoa){
  let n = 0;
  for(const r of (aoa || []).slice(0,120)){
    for(const v of (r || [])){
      const t = text(v);
      if(t && /[가-힣A-Za-z]/.test(t)) n++;
    }
  }
  return n;
}

function aoaHasSalaryHeaderLikeText(aoa){
  return aoaScoreForHeaders(aoa) >= 2;
}
function aoaScoreForHeaders(aoa){
  const joined = (aoa || []).slice(0,80).map(r => (r||[]).map(text).join(' ')).join(' ');
  let score = 0;
  if(/직\s*명|직명/.test(joined)) score++;
  if(/본봉|기본급/.test(joined)) score++;
  if(/지급액\s*계|지급액계|지급총액|총지급액/.test(joined)) score++;
  if(/소계\s*\(?교원\)?|소계교원/.test(joined)) score++;
  return score;
}
function sheetXmlToAoaByRegex(raw, sharedStrings){
  const aoa = [];
  const rowRe = /<(?:\w+:)?row\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?row>/g;
  let rm, implicitRow = 0;
  while((rm = rowRe.exec(raw))){
    const rowAttrs = rm[1] || '';
    const body = rm[2] || '';
    const rIdx = Number(getXmlAttr(rowAttrs, 'r') || (++implicitRow)) - 1;
    implicitRow = Math.max(implicitRow, rIdx + 1);
    if(!aoa[rIdx]) aoa[rIdx] = [];
    const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let cm;
    while((cm = cellRe.exec(body))){
      const attrs = cm[1] || '';
      const cbody = cm[2] || '';
      const ref = getXmlAttr(attrs, 'r') || '';
      const cIdx = ref ? colRefToIndex(ref) : aoa[rIdx].length;
      aoa[rIdx][cIdx] = cellValueFromXml(attrs, cbody, sharedStrings);
    }
  }
  return aoa.map(r => r || []);
}
function getXmlAttr(attrs, name){
  const m = String(attrs || '').match(new RegExp('(?:^|\\s)' + name + '="([^"]*)"'));
  return m ? xmlDecode(m[1]) : '';
}
function cellValueFromXml(attrs, body, sharedStrings){
  const t = getXmlAttr(attrs, 't');
  if(t === 'inlineStr'){
    const parts = [];
    const tRe = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
    let tm; while((tm = tRe.exec(body))) parts.push(xmlDecode(tm[1]));
    return parts.join('');
  }
  const vm = body.match(/<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/);
  const raw = vm ? xmlDecode(vm[1]) : '';
  if(t === 's') return sharedStrings[Number(raw)] ?? '';
  if(t === 'str') return raw;
  if(raw !== ''){
    const n = Number(raw);
    return Number.isFinite(n) && /^-?\d+(\.\d+)?$/.test(raw) ? n : raw;
  }
  const fm = body.match(/<(?:\w+:)?f\b[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
  return fm ? '=' + xmlDecode(fm[1]) : '';
}
function xmlDecode(s){
  return String(s ?? '')
    .replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"')
    .replace(/&apos;/g,"'").replace(/&amp;/g,'&');
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
  let score = 0;
  const n = norm(name);
  const isSalary = /(교직원.*보수|보수.*일람|보수|급여|봉급|인건비|교직원)/.test(name);
  if(!isSalary) return 0;
  if(/교직원보수일람표/.test(n)) score += 120;
  if(/보수/.test(n)) score += 50;
  if(/급여|봉급|인건비/.test(n)) score += 25;
  if(/\(년\)|년간|연간|연/.test(n)) score += 35;
  if(/\(월\)|월간|월급여|월/.test(n)) score -= 25;
  if(/비월정|간이세액|세출|세입|퇴직|보험|명시이월|사고이월/.test(name)) score -= 120;
  return Math.max(0, score);
}
function effectiveMaxCols(aoa, maxRows=100){
  // 일부 한컴/HCell XLSX는 실제 데이터가 A~T까지만 있어도 !ref가 A1:XFB41처럼 저장됩니다.
  // 이 범위를 그대로 돌면 헤더 탐색이 느려지거나 빈 열에 밀려 실패하므로 실제 값이 있는 마지막 열까지만 사용합니다.
  let last = 0;
  for(let r=0; r<Math.min(maxRows, aoa.length); r++){
    const row = aoa[r] || [];
    const limit = Math.min(row.length || 0, 300);
    for(let c=0; c<limit; c++){
      if(text(row[c])) last = Math.max(last, c);
    }
  }
  return Math.min(Math.max(last + 6, 25), 120);
}

function parseSalarySheet(sheetName, aoa){
  const preview = aoa.slice(0,60).map((r,i)=>({행:i+1, 내용:r.map(text).filter(Boolean).join(' | ').slice(0,300)})).filter(x=>x.내용);
  if(!aoa.length) return {ok:false, sheetName, message:'빈 시트', preview};
  const maxRows = Math.min(100, aoa.length), maxCols = effectiveMaxCols(aoa, maxRows);
  const detected = detectSalaryHeader(aoa, maxRows, maxCols);
  if(!detected.ok){
    return {ok:false, sheetName, message:detected.message, preview};
  }
  const {headerRow, jobCol, nameCol, baseCol, totalCol} = detected;
  const allowanceCols = [];
  for(let c=baseCol+1; c<totalCol; c++){
    const h = headerNameForCol(aoa, c, headerRow) || cleanHeader(columnText(aoa, c, maxRows));
    if(h && !/성\s*명|호봉|직명|일련|번호|주민등록|본봉|기본급|지급액계|지급액.*계|합계|소계|소득세|주민세|본인부담|공제액계|실수령액|사학연금|국민연금|사회보험/.test(h)){
      allowanceCols.push({ col:c, name:h || `열${c+1}` });
    }
  }
  const teacherEnd = findRow(aoa, headerRow+1, /소계\s*\(?교원\)?|소계교원/);
  const staffEnd = findRow(aoa, (teacherEnd>=0?teacherEnd+1:headerRow+1), /소계\s*\(?(직원|일반직)\)?|소계직원|소계일반직/);
  const teacherRows = rowsInRange(aoa, headerRow+1, teacherEnd>=0?teacherEnd:staffEnd, {jobCol,nameCol,baseCol,totalCol,allowanceCols});
  const staffRows = rowsInRange(aoa, teacherEnd>=0?teacherEnd+1:headerRow+1, staffEnd, {jobCol,nameCol,baseCol,totalCol,allowanceCols});
  const teacherSubtotal = teacherEnd>=0 ? parseDataRow(aoa[teacherEnd], {jobCol,nameCol,baseCol,totalCol,allowanceCols}) : null;
  const staffSubtotal = staffEnd>=0 ? parseDataRow(aoa[staffEnd], {jobCol,nameCol,baseCol,totalCol,allowanceCols}) : null;
  if(!teacherRows.length && !staffRows.length && !teacherSubtotal && !staffSubtotal){
    return {ok:false, sheetName, message:'헤더는 찾았으나 교원/직원 데이터 범위를 읽지 못했습니다.', preview};
  }
  const summary = summarizeSalary(teacherRows, staffRows, teacherSubtotal, staffSubtotal, allowanceCols);
  return {ok:true, sheetName, header:{headerRow:headerRow+1, jobCol:jobCol+1, nameCol:nameCol+1, baseCol:baseCol+1, totalCol:totalCol+1, allowanceCols:allowanceCols.map(x=>({열:x.col+1, 이름:x.name}))}, ranges:{teacherStart:headerRow+2, teacherEnd:teacherEnd>=0?teacherEnd+1: null, staffStart:teacherEnd>=0?teacherEnd+2:headerRow+2, staffEnd:staffEnd>=0?staffEnd+1:null}, summary, teacherRows, staffRows, teacherSubtotal, staffSubtotal, preview};
}
function detectSalaryHeader(aoa, maxRows, maxCols){
  const direct = detectSalaryHeaderDirect(aoa, maxRows, maxCols);
  if(direct.ok) return direct;
  let best = null;
  for(let r=0; r<maxRows; r++){
    const rowWindow = [];
    for(let c=0; c<maxCols; c++){
      rowWindow[c] = [aoa[r-1]?.[c], aoa[r]?.[c], aoa[r+1]?.[c]].map(text).filter(Boolean).join(' ');
    }
    const jobCol = findColInTexts(rowWindow, [/^직명$/, /직\s*명/]);
    const nameCol = findColInTexts(rowWindow, [/^성\s*명$/, /성\s*명/]);
    const baseCol = findColInTexts(rowWindow, [/본봉/, /기본급/]);
    const totalCol = findColInTexts(rowWindow, [/지급액계/, /지급액\s*계/, /월지급액계/, /지급총액/, /총지급액/]);
    let score = 0;
    if(jobCol>=0) score+=4; if(baseCol>=0) score+=4; if(totalCol>=0) score+=4; if(nameCol>=0) score+=1;
    if(baseCol>=0 && totalCol>baseCol) score+=3;
    const all = rowWindow.join(' ');
    if(/교직원|보수|일람/.test(all)) score-=1;
    if(score > (best?.score ?? -1)) best = {score, headerRow:r, jobCol, nameCol, baseCol, totalCol};
  }
  // 열 전체 검색 보조: 병합/다중행/한컴 XLSX에서 헤더가 서로 다른 행에 흩어져 있을 때 사용
  const colText = Array.from({length:maxCols},(_,c)=>columnText(aoa,c,maxRows));
  const colDetected = {
    headerRow: findHeaderRow(aoa, []),
    jobCol: findColInTexts(colText, [/^직명$/, /직\s*명/]),
    nameCol: findColInTexts(colText, [/^성\s*명$/, /성\s*명/]),
    baseCol: findColInTexts(colText, [/본봉/, /기본급/]),
    totalCol: findColInTexts(colText, [/지급액계/, /지급액\s*계/, /월지급액계/, /지급총액/, /총지급액/]),
    score: 0
  };
  if(colDetected.jobCol>=0) colDetected.score+=3;
  if(colDetected.baseCol>=0) colDetected.score+=3;
  if(colDetected.totalCol>=0) colDetected.score+=3;
  if(colDetected.baseCol>=0 && colDetected.totalCol>colDetected.baseCol) colDetected.score+=2;
  const picked = (best && best.score >= colDetected.score) ? best : colDetected;
  if(picked.jobCol < 0 || picked.baseCol < 0 || picked.totalCol < 0 || picked.totalCol <= picked.baseCol){
    return {ok:false, message:`헤더 미발견: 직명=${picked.jobCol+1}, 본봉=${picked.baseCol+1}, 지급액계=${picked.totalCol+1}`};
  }
  if(picked.headerRow == null || picked.headerRow < 0){
    picked.headerRow = findHeaderRow(aoa, [picked.jobCol, picked.baseCol, picked.totalCol]);
  }
  return {ok:true, ...picked};
}

function detectSalaryHeaderDirect(aoa, maxRows, maxCols){
  // 보수표는 기관마다 병합/다중행 헤더가 달라도 보통 한 행 안에 직명·본봉·지급액계가 존재합니다.
  // 먼저 상단 80행 전체에서 각 셀 단위로 가장 직접적인 헤더를 찾습니다.
  const cols = {jobCol:-1, nameCol:-1, baseCol:-1, totalCol:-1};
  let headerRow = -1;
  const scanRows = Math.min(maxRows, 80);
  for(let r=0; r<scanRows; r++){
    for(let c=0; c<maxCols; c++){
      const v = norm(text(aoa[r]?.[c]));
      if(!v) continue;
      if(cols.jobCol<0 && /^직명$|직\s*명/.test(v)){ cols.jobCol=c; headerRow = headerRow<0?r:Math.max(headerRow,r); }
      if(cols.nameCol<0 && /^성명$|성\s*명/.test(v)){ cols.nameCol=c; }
      if(cols.baseCol<0 && /본봉|기본급/.test(v)){ cols.baseCol=c; headerRow = headerRow<0?r:Math.max(headerRow,r); }
      if(cols.totalCol<0 && /(지급액계|지급액\s*계|월지급액계|지급총액|총지급액)/.test(v)){ cols.totalCol=c; headerRow = headerRow<0?r:Math.max(headerRow,r); }
    }
  }
  // 어떤 파일은 '지급액계'가 아닌 '지급액 계'가 줄바꿈으로 분리됩니다. 열 전체 텍스트로 한 번 더 찾습니다.
  const colText = Array.from({length:maxCols},(_,c)=>norm(columnText(aoa,c,scanRows)));
  if(cols.jobCol<0) cols.jobCol = findColInTexts(colText, [/^직명$/, /직\s*명/]);
  if(cols.nameCol<0) cols.nameCol = findColInTexts(colText, [/^성명$/, /성\s*명/]);
  if(cols.baseCol<0) cols.baseCol = findColInTexts(colText, [/본봉/, /기본급/]);
  if(cols.totalCol<0) cols.totalCol = findColInTexts(colText, [/지급액계/, /지급액\s*계/, /월지급액계/, /지급총액/, /총지급액/, /지급총액/, /총지급액/]);
  if(headerRow<0) headerRow = findHeaderRow(aoa, [cols.jobCol, cols.baseCol, cols.totalCol]);
  if(cols.jobCol>=0 && cols.baseCol>=0 && cols.totalCol>cols.baseCol){
    return {ok:true, headerRow, jobCol:cols.jobCol, nameCol:cols.nameCol, baseCol:cols.baseCol, totalCol:cols.totalCol, score:99};
  }
  return {ok:false};
}

function columnText(aoa, c, maxRows=80){
  return aoa.slice(0,maxRows).map(r=>r?.[c]).map(text).filter(Boolean).join(' ');
}
function findColInTexts(texts, patterns){
  let best=-1, bestScore=-1;
  texts.forEach((t,i)=>{
    const n=norm(t);
    patterns.forEach((re,idx)=>{
      if(re.test(n) || re.test(t)){
        const sc=100-idx;
        if(sc>bestScore){best=i; bestScore=sc;}
      }
    });
  });
  return best;
}
function findCol(colText, patterns){ return findColInTexts(colText, patterns); }
function findHeaderRow(aoa, cols){
  let best=0, score=-1;
  aoa.slice(0,100).forEach((r,i)=>{
    let s=0;
    if(cols && cols.length) cols.forEach(c=>{ if(text(r?.[c])) s++; });
    const all=r.map(text).join(' ');
    if(/직\s*명/.test(all))s+=3;
    if(/본봉|기본급/.test(all))s+=3;
    if(/지급액계|지급액\s*계|지급총액|총지급액/.test(all))s+=3;
    if(/소득세|실수령액|공제액/.test(all))s+=1;
    if(s>score){score=s;best=i;}
  });
  return best;
}
function findRow(aoa, start, re){ for(let r=Math.max(0,start); r<aoa.length; r++){ if(re.test(aoa[r].map(text).join(' '))) return r; } return -1; }
function rowsInRange(aoa, start, end, cfg){
  const last = end>=0 ? end : aoa.length;
  const rows=[];
  for(let r=Math.max(0,start); r<last; r++){
    const row = parseDataRow(aoa[r], cfg);
    if(row && (row.직명 || row.성명 || row.본봉 || row.지급액계 || Object.values(row.수당||{}).some(Boolean))) rows.push({...row, 행:r+1});
  }
  return rows.filter(r=> !/소계|합계|계\s*$/.test(`${r.직명} ${r.성명}`) && (r.직명 || r.성명));
}
function parseDataRow(row, cfg){
  if(!row) return null;
  const out = {직명:text(row[cfg.jobCol]), 성명:cfg.nameCol>=0?text(row[cfg.nameCol]):'', 본봉:toNum(row[cfg.baseCol]), 지급액계:toNum(row[cfg.totalCol])};
  out.수당 = {};
  cfg.allowanceCols.forEach(a=>{ out.수당[a.name]=toNum(row[a.col]); });
  return out;
}
function summarizeSalary(teacherRows, staffRows, teacherSubtotal, staffSubtotal, allowanceCols){
  const rowHasMoney = r => (r.본봉 || r.지급액계 || Object.values(r.수당||{}).some(Boolean));
  const sumRows = (rows) => ({인원:rows.filter(rowHasMoney).length, 본봉:rows.reduce((s,r)=>s+r.본봉,0), 지급액계:rows.reduce((s,r)=>s+r.지급액계,0)});
  const t = sumRows(teacherRows), s = sumRows(staffRows);
  const allowances = allowanceCols.map(a=>({항목:a.name, 교원금액:teacherRows.reduce((sum,r)=>sum+(r.수당[a.name]||0),0), 직원금액:staffRows.reduce((sum,r)=>sum+(r.수당[a.name]||0),0), 교원인원:teacherRows.filter(r=>(r.수당[a.name]||0)>0).length, 직원인원:staffRows.filter(r=>(r.수당[a.name]||0)>0).length})).filter(x=>x.교원금액||x.직원금액);
  return {교원:{...t, 소계본봉:teacherSubtotal?.본봉||0, 소계지급액계:teacherSubtotal?.지급액계||0}, 직원:{...s, 소계본봉:staffSubtotal?.본봉||0, 소계지급액계:staffSubtotal?.지급액계||0}, 수당:allowances};
}
function cleanHeader(s){ return text(s).replace(/\([A-Z][^)]*\)/g,'').replace(/\s+/g,' ').trim().slice(0,40); }
function headerNameForCol(aoa, col, headerRow){
  const start = Math.max(0, headerRow - 2);
  const end = Math.min(aoa.length - 1, headerRow + 1);
  const vals = [];
  for(let r=start; r<=end; r++){
    const v = text(aoa[r]?.[col]);
    if(!v) continue;
    if(toNum(v) || /소계|합계|원장|교사|직원|성명|직명|호봉|순번|번호|구분|주민등록|소득세|주민세|공제|실수령/.test(v)) continue;
    vals.push(v);
  }
  return cleanHeader([...new Set(vals)].join(' '));
}
function parseRetireSheet(sheetName, aoa){
  // v42: 퇴직 적립 여부는 `적립금이월액` 구역의 `계` 행 금액을 기준으로 판단합니다.
  // 금액 일치 여부는 보지 않고, 이 값이 양수인지 여부만 이후 PDF 편성 여부와 비교합니다.
  const maxRows = Math.min(80, aoa.length);
  const maxCols = effectiveMaxCols(aoa, maxRows);
  const cells = [];
  for(let r=0; r<maxRows; r++){
    for(let c=0; c<maxCols; c++){
      const v = text(aoa[r]?.[c]);
      if(v) cells.push({r,c,v,n:norm(v)});
    }
  }

  // 1) `적립금이월액` 헤더가 병합되어 있으면 그 아래/주변 열 중 `계` 행의 금액을 읽습니다.
  const carryHeaders = cells.filter(x => /적립금.*이월액|이월.*적립금|적립금이월액/.test(x.n));
  const totalRows = cells.filter(x => /^(계|합계|총계)$/.test(x.n));
  const picked = [];
  for(const h of carryHeaders){
    // 병합 헤더는 보통 해당 구역의 왼쪽 끝 셀에만 값이 있고, 실제 금액은 오른쪽 0~4열에 있습니다.
    const cStart = h.c;
    const cEnd = Math.min(maxCols - 1, h.c + 5);
    for(const tr of totalRows){
      if(tr.r <= h.r) continue;
      const row = aoa[tr.r] || [];
      for(let c=cStart; c<=cEnd; c++){
        const n = toNum(row[c]);
        if(n > 0){
          picked.push({행:tr.r+1, 열:c+1, 값:n, 헤더:h.v, 주변:(row.map(text).filter(Boolean).join(' ')).slice(0,180)});
        }
      }
    }
  }

  // 2) 어떤 서식은 `적립금이월액` 텍스트가 공유문자열/병합셀 때문에 복원되지 않을 수 있습니다.
  // 이 경우에는 퇴직금적립현황 표에서 `계` 행의 가장 오른쪽 양수 금액을 보조 기준으로 사용합니다.
  // 단, `퇴직` 관련 시트에서만 이 함수가 호출되므로 급여표 숫자와 혼동하지 않습니다.
  if(!picked.length){
    for(const tr of totalRows){
      const row = aoa[tr.r] || [];
      const nums = [];
      for(let c=0; c<Math.min(row.length, maxCols); c++){
        const n = toNum(row[c]);
        if(n > 0) nums.push({행:tr.r+1, 열:c+1, 값:n, 헤더:'계 행 우측 금액', 주변:(row.map(text).filter(Boolean).join(' ')).slice(0,180)});
      }
      if(nums.length){
        picked.push(nums[nums.length-1]);
        break;
      }
    }
  }

  // 같은 행/열 중복 제거 후, `계` 기준은 하나의 대표값만 사용합니다.
  const unique = [];
  const seen = new Set();
  for(const c of picked){
    const key = `${c.행}:${c.열}:${c.값}`;
    if(!seen.has(key)){ seen.add(key); unique.push(c); }
  }
  const representative = unique.length ? unique[0] : null;
  const positive = representative ? Number(representative.값 || 0) : 0;
  return {
    sheetName,
    positiveAmount: positive,
    hasRetirementAmount: positive > 0,
    retireColumns: carryHeaders.map(h=>({열:h.c+1, 헤더:h.v})),
    positiveCells: representative ? [representative] : [],
    ignoredPositiveCells: unique.slice(1,10),
    기준:'적립금이월액의 계'
  };
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
      out.push({페이지:l.page, 행:i+1, 목:totalRow.항목, 상위항목:'', 항목:totalRow.항목, PDF금액:totalRow.금액, PDF금액천원:Math.round(totalRow.금액/1000), 보조금금액:totalRow.보조금, 수익자금액:totalRow.수익자, 기타금액:totalRow.기타, 인원:null, 산출기초:'', 구분:'총액행'});
      continue;
    }
    const calc = parseCalcLine(t);
    if(calc){
      const name = calc.name || findCalcName(lines, i) || currentTotal?.항목 || '산출항목미상';
      out.push({페이지:l.page, 행:i+1, 목:currentTotal?.항목 || '', 상위항목:currentTotal?.항목 || '', 항목:name, PDF금액:calc.amount, PDF금액천원:Math.round(calc.amount/1000), 인원:calc.people, 누락항목:calc.missing || [], 산출기초:t, 구분:'산출기초'});
    }
  }
  return out.filter(x=>!LEGAL_RE.test(x.항목));
}
function parseTotalBudgetRow(t){
  if(/산출|예산구분|발행일|과\s*목|보조금\s*및|수익자\s*부담|예산액\s*전년도|비교\s*증감/.test(t)) return null;
  const raw = text(t).replace(/−/g,'-').replace(/\s+/g,' ').trim();
  if(!raw) return null;

  // ① 행복한아이들 등: 보조금/수익자/기타/합계 4개 금액 열
  let m = raw.match(/^\s*([^0-9=]{2,60}?)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s*$/);
  if(m){
    const name = normalizeMokName(text(m[1]));
    if(!validBudgetRowName(name)) return null;
    return {항목:name, 보조금:toNum(m[2])*1000, 수익자:toNum(m[3])*1000, 기타:toNum(m[4])*1000, 금액:toNum(m[5])*1000, 형식:'재원별'};
  }

  // ② 신동탄재크와콩나무 등: 예산액/전년도예산액/비교증감 3개 금액 열
  // 이 형식에서는 첫 번째 숫자가 현재 예산액입니다.
  m = raw.match(/^\s*([^0-9=]{2,80}?)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s+(-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+)\s*$/);
  if(m){
    const name = normalizeMokName(text(m[1]));
    if(!validBudgetRowName(name)) return null;
    return {항목:name, 보조금:0, 수익자:0, 기타:0, 금액:toNum(m[2])*1000, 전년도금액:toNum(m[3])*1000, 비교증감:toNum(m[4])*1000, 형식:'예산액'};
  }
  return null;
}
function normalizeMokName(s){
  return text(s).replace(/\s+/g,'').replace(/^[|:：·ㆍ]+|[|:：·ㆍ]+$/g,'').trim();
}
function validBudgetRowName(name){
  if(!name) return false;
  if(name.length < 2 || name.length > 35) return false;
  if(/본예산|단위|천원|예산구분|회계연도|학교명|세입|세출|총괄표|김담은|경기도교육청/.test(name)) return false;
  if(/^[\d,\-]+$/.test(name)) return false;
  return true;
}
function extractCalcItemNameFromLine(t){
  let s = text(t || '');
  s = s.replace(/\(본예산\)|\(보조금및지원금\)|\(수익자부담금\)|\(그밖의수입\)/g, ' ')
       .replace(/\s+/g, ' ')
       .trim();
  if(!s) return '';
  // 산출내역과 산출기초가 같은 줄에 붙어 있는 PDF를 처리합니다.
  // 예: "직책수당 (원장500,000원+부원장200,000원+교사50,000원)*12월 = 9,000,000"
  // 예: "[방]근속수당 20,000원*1명*12월 = 240,000"
  let m = s.match(/^(.{2,45}?)(?=\s*(?:[\{\(]\s*[^=]{0,20}?[0-9,]+원|[0-9,]+원))/);
  if(!m) return '';
  let name = cleanCalcNameCandidate(m[1]);
  if(!name) return '';
  return name;
}
function parseCalcLine(t){
  const extractedName = extractCalcItemNameFromLine(t);
  const compact = t.replace(/\s+/g,'').replace(/＝/g,'=');
  let m = compact.match(/([0-9,]+)원\*([0-9,]+)명\*([0-9,]+)(?:월|개월|개?월)=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:toNum(m[2]), months:toNum(m[3]), amount:toNum(m[4]), missing:[], name:extractedName};

  // 인원 수가 빠진 형태: 3,600,000원*12월=43,200,000
  m = compact.match(/([0-9,]+)원\*([0-9,]+)(?:월|개월|개?월)=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:null, months:toNum(m[2]), amount:toNum(m[3]), missing:['인원 수'], name:extractedName};

  // 월수/개월 수가 빠진 형태: 100,000원*5명=500,000
  m = compact.match(/([0-9,]+)원\*([0-9,]+)명=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:toNum(m[2]), months:null, amount:toNum(m[3]), missing:['월수'], name:extractedName};

  // 산출식은 있으나 인원/월수 구조가 불완전한 형태도 금액은 잡아 둡니다.
  if(/원/.test(compact) && /=/.test(compact)){
    const amountMatch = compact.match(/=([0-9,]+)$/);
    if(amountMatch){
      const peopleMatch = compact.match(/([0-9,]+)명/);
      const monthMatch = compact.match(/([0-9,]+)(?:월|개월|개?월)/);
      const unitMatch = compact.match(/([0-9,]+)원/);
      const missing = [];
      if(!peopleMatch) missing.push('인원 수');
      if(!monthMatch) missing.push('월수');
      return {
        unit: unitMatch ? toNum(unitMatch[1]) : null,
        people: peopleMatch ? toNum(peopleMatch[1]) : null,
        months: monthMatch ? toNum(monthMatch[1]) : null,
        amount: toNum(amountMatch[1]),
        missing,
        name: extractedName
      };
    }
  }
  return null;
}

function isLaborBasisTarget(x){
  const s = norm((x?.목 || '') + ' ' + (x?.항목 || '') + ' ' + (x?.산출기초 || ''));
  if(LEGAL_RE.test(s)) return false;
  return /(교원급여|교원수당|직원급여|직원수당|그밖의인건비|인건비|급여|수당|상여|휴가비|정액급식|급식보조|보조비|처우개선|연구|직급|직책|자가운전)/.test(s);
}
function addBasisIssues(issues, calcs){
  const seen = new Set();
  for(const x of calcs || []){
    if(!isLaborBasisTarget(x)) continue;
    const missing = Array.isArray(x.누락항목) ? x.누락항목 : (Array.isArray(x.missing) ? x.missing : []);
    if(!missing.length) continue;
    const item = cleanItemName(x.항목 || x.목 || '해당 항목');
    const msg = `${item} 산출기초에 ${missing.join(', ')} 미표시`;
    const basis = `PDF ${x.페이지 || ''}쪽 ${x.목 ? `목 ${x.목}, ` : ''}산출기초: ${x.산출기초 || ''}`;
    const key = msg + '|' + basis;
    if(seen.has(key)) continue;
    seen.add(key);
    issues.push({번호:issues.length+1, 지적내용:msg, 근거:basis});
  }
}
function cleanCalcNameCandidate(raw){
  let s = text(raw || '');
  s = s.replace(/\(본예산\)|\(보조금및지원금\)|\(수익자부담금\)|\(그밖의수입\)/g,' ')
       .replace(/교\s*$/,'').replace(/환\s*$/,'').replace(/운\s*$/,'').replace(/경\s*$/,'')
       .replace(/\s+/g,' ').trim();
  if(!s) return '';
  if(/[0-9,]+원\s*\*/.test(s)) return '';
  if(/예산구분|발행일|산출|과\s*목|보조금|수익자|합계/.test(s)) return '';
  const cleaned = s.replace(/\s+/g,'');
  if(cleaned.length < 2 || cleaned.length > 40) return '';
  if(/^\d+$/.test(cleaned)) return '';
  return cleaned;
}
function findCalcName(lines, idx){
  // 산출항목명은 대부분 산출기초 바로 위에 있습니다.
  // 예: "근속수당" 다음 줄에 "{(20,000원*2명)+...}=960,000" 형태.
  // 이전 버전은 아래쪽 후보를 먼저 보아 다음 항목명([방]근속수당 등)과 잘못 매칭되는 경우가 있었습니다.
  const offsets = [-1,-2,-3,-4,-5,-6,1,2,3,4,5,6];
  for(const off of offsets){
    const j = idx + off;
    if(j < 0 || j >= lines.length) continue;
    const raw = lines[j]?.text || '';
    // 총액행 또는 산출기초식이 들어 있는 줄은 항목명 후보에서 제외합니다.
    if(parseTotalBudgetRow(raw)) continue;
    if(/[0-9,]+\s*원|=\s*[0-9,]+/.test(text(raw))) continue;
    const cand = cleanCalcNameCandidate(raw);
    if(cand) return cand;
  }
  return '';
}


function findOffMokMatches(calcs, expectedMok, diff, category){
  const absDiff = Math.abs(Number(diff||0));
  if(absDiff <= 1000) return [];
  const target = norm(expectedMok);
  const categoryNorm = norm(category || expectedMok);
  const isPay = /급여/.test(category);
  const isAllowance = /수당/.test(category);
  return calcs.filter(x=>{
    if(!x || !x.PDF금액) return false;
    if(LEGAL_RE.test(x.항목 || '')) return false;
    const mok = norm(x.목 || x.상위항목 || '');
    if(!mok || mok.includes(target)) return false;
    if(!closeMoney(Math.abs(x.PDF금액), absDiff)) return false;
    const label = norm((x.항목 || '') + ' ' + (x.산출기초 || ''));
    if(isPay) return /(급여|인건비|보수)/.test(label);
    if(isAllowance) return /(수당|상여|휴가비|식대|정액급식|자가운전|직급|직책|연구|관리업무)/.test(label);
    return true;
  });
}
function addOffMokDiffIssues(issues, calcs, checks){
  const seen = new Set(issues.map(i=>i.지적내용 + '|' + i.근거));
  for(const c of checks){
    const diff = amountDiff(c.excelAmount, c.pdfAmount);
    const matches = findOffMokMatches(calcs, c.expectedMok, diff, c.category);
    for(const x of matches){
      const kind = /교원/.test(c.category) ? '교원' : /직원/.test(c.category) ? '직원' : '해당';
      const text = `${cleanItemName(x.항목)} ${shortWon(x.PDF금액)}을 ${c.expectedMok} 목이 아닌 ${x.목 || x.상위항목} 목의 산출내역에 편성`;
      const basis = `${c.category} 차액 ${fmt(Math.abs(diff))}과 일치. PDF ${x.페이지}쪽: ${x.산출기초}`;
      const key = text + '|' + basis;
      if(seen.has(key)) continue;
      seen.add(key);
      issues.push({번호:issues.length+1, 지적내용:text, 근거:basis});
    }
  }
}


function exactMokCalcs(calcs, mokName){
  const target = norm(mokName);
  return calcs.filter(x=>norm(x.목 || x.상위항목 || '') === target);
}
function sumAmount(items){ return items.reduce((s,x)=>s+Number(x.PDF금액||0),0); }
function sumPeople(items){ return items.reduce((s,x)=>s+Number(x.인원||0),0); }
function pdfPeopleKey(x){
  // 보조금/수익자부담금처럼 재원이 나뉜 같은 항목은 같은 인원으로 보아 중복 계산하지 않습니다.
  // 다만 방과후교원급여처럼 별도 인력군은 교원급여와 분리합니다.
  const raw = norm((x?.항목 || '') + ' ' + (x?.산출기초 || ''));
  if(/방과후.*교원|방과후.*교사|방과후교원급여/.test(raw)) return '방과후교원급여';
  if(/교원급여|교사급여|원장급여/.test(raw)) return '교원급여';
  if(/사무.*급여/.test(raw)) return '사무직원급여';
  if(/조리.*급여/.test(raw)) return '조리직원급여';
  if(/환경.*급여|미화.*급여/.test(raw)) return '환경미화원급여';
  if(/보조교사.*급여/.test(raw)) return '보조교사급여';
  if(/차량기사.*급여|량기사.*급여/.test(raw)) return '차량기사급여';
  if(/차량보조.*급여/.test(raw)) return '차량보조급여';
  if(/영양사.*급여|영양사.*인건비/.test(raw)) return '영양사급여';
  return allowanceKey(x?.항목 || '') || norm(x?.항목 || x?.목 || '');
}

function cleanItemName(name){
  let s = String(name || '').trim();
  if(/^량기사/.test(s)) s = '차' + s;
  if(/^량보조/.test(s)) s = '차' + s;
  return s;
}

function sumPeopleDistinct(items){
  const byName = new Map();
  for(const x of items){
    const name = pdfPeopleKey(x);
    if(!name) continue;
    const people = Number(x.인원 || 0);
    if(!people) continue;
    byName.set(name, Math.max(byName.get(name) || 0, people));
  }
  return [...byName.values()].reduce((a,b)=>a+b,0);
}
function detailVerdict(label, excelAmount, pdfAmount, excelPeople, pdfPeople, offMatches=[]){
  const parts=[];
  const ea=Number(excelAmount||0), pa=Number(pdfAmount||0);
  const ep=Number(excelPeople||0), pp=Number(pdfPeople||0);
  if(!closeMoney(ea, pa)){
    const diff=ea-pa;
    if(offMatches && offMatches.length){
      const m=offMatches[0];
      parts.push(`금액 차이(${label} ${shortWon(Math.abs(diff))} 차이 → ${m.목 || m.상위항목} 목에 ${cleanItemName(m.항목)} ${fmt(m.PDF금액)} 편성)`);
    }else{
      parts.push(`금액 차이(엑셀 ${fmt(ea)} / PDF ${fmt(pa)} / 차이 ${fmt(Math.abs(diff))})`);
    }
  }
  // v41: 인원 차이는 검토 대상에서 제외합니다.
  return parts.length ? parts.join(' / ') : '일치';
}
function totalRowByName(totals, name){ return totals.find(x=>norm(x.항목)===norm(name)); }
function allowanceGroupCandidates(groupRow, kind){
  if(!groupRow) return [];
  const out=[];
  if(groupRow.PDF금액) out.push({label:`${kind}수당`, amount:Number(groupRow.PDF금액||0)});
  if(groupRow.수익자금액) out.push({label:`${kind}수당`, amount:Number(groupRow.수익자금액||0)});
  if(groupRow.보조금금액) out.push({label:`${kind}수당`, amount:Number(groupRow.보조금금액||0)});
  if(groupRow.기타금액) out.push({label:`${kind}수당`, amount:Number(groupRow.기타금액||0)});
  return out.filter((x,i,a)=>x.amount && a.findIndex(y=>y.amount===x.amount)===i);
}
function findAllowancePdfMatches(kind, allowanceName, amount, pdfItems){
  const key = allowanceKey(allowanceName);
  return pdfItems.filter(x=>{
    if(x.구분 !== '산출기초') return false;
    if(!allowancePdfHit(kind, key, amount, x)) return false;
    // 같은 명칭이어도 PDF 추출이 엇갈린 경우를 막기 위해 금액도 함께 확인합니다.
    // 다만 PDF가 같은 항목을 재원별로 나눈 경우에는 합산 단계에서 처리되므로
    // 개별 산출기초 금액이 엑셀금액 이하인 경우까지 허용합니다.
    const pdfAmt = Number(x.PDF금액 || 0);
    const excelAmt = Number(amount || 0);
    return !excelAmt || closeMoney(pdfAmt, excelAmt) || pdfAmt <= excelAmt;
  });
}
function integratedText(kind, list, amountField, groupLabel, groupAmount){
  const expr = list.map(a=>`${a.항목}(${fmt(a[amountField])})`).join(' + ');
  return `${expr} = ${groupLabel}(${fmt(groupAmount)})으로 통합편성 추정`;
}


function roleTokenFromPdfItemName(name){
  const n = norm(name);
  const tokens = ['방과후교원','방과후교사','차량기사','차량보조','보조교사','사무직원','조리직원','환경미화','영양사','교원','교사','원장','직원'];
  return tokens.find(t => n.includes(norm(t))) || '';
}
function excelBaseByRole(rows, token){
  const tk = norm(token);
  if(!tk || !rows) return {amount:0, people:0};
  const matched = rows.filter(r => {
    const hay = norm(`${r.직명||''} ${r.성명||''}`);
    if(tk.includes('교원') && (hay.includes('원장') || hay.includes('교사'))) return true;
    return hay.includes(tk.replace('직원','')) || hay.includes(tk);
  });
  return { amount: matched.reduce((s,r)=>s+Number(r.본봉||0),0), people: matched.filter(r=>Number(r.본봉||0)>0).length };
}
function addPayBreakdownRows(rows, kind, salaryObj, payCalcs, allCalcs){
  if(!salaryObj || !payCalcs || !payCalcs.length) return;
  const groups = new Map();
  for(const x of payCalcs){
    const token = roleTokenFromPdfItemName(x.항목);
    if(!token || token === '교원' || token === '직원') continue;
    const key = x.항목;
    if(!groups.has(key)) groups.set(key, []);
    groups.get(key).push(x);
  }
  for(const [name, items] of groups){
    const token = roleTokenFromPdfItemName(name);
    const ex = excelBaseByRole(kind === '교원' ? salaryObj.teacherRows : salaryObj.staffRows, token);
    const pdfAmount = sumAmount(items);
    const pdfPeople = sumPeopleDistinct(items) || sumPeople(items) || 0;
    if(!ex.amount && !pdfAmount) continue;
    const expectedMok = kind === '교원' ? '교원급여' : '직원급여';
    const off = findOffMokMatches(allCalcs || [], expectedMok, amountDiff(ex.amount, pdfAmount), name);
    rows.push({구분:'급여', 항목:name, 엑셀금액:ex.amount, 엑셀인원:ex.people, PDF금액:pdfAmount, PDF인원:pdfPeople, 검토결과:detailVerdict(name, ex.amount, pdfAmount, ex.people, pdfPeople, off)});
  }
}

function isRetirementName(s){
  const n = norm(s || '');
  return /퇴직/.test(n) && /(적립|퇴직금|충당|급여)/.test(n);
}
function getExcelRetirementSummary(retirementSheets){
  const sheets = retirementSheets || [];
  const amount = sheets.reduce((sum,r)=>sum+Number(r.positiveAmount||0),0);
  return {
    amount,
    has: amount > 0,
    sheets: sheets.filter(r=>Number(r.positiveAmount||0)>0).map(r=>r.sheetName),
    cells: sheets.flatMap(r=>(r.positiveCells||[]).map(c=>({...c, sheetName:r.sheetName}))).slice(0,30)
  };
}
function getPdfRetirementSummary(pdfItems){
  const items = pdfItems || [];
  const totalRows = items.filter(x=>x.구분==='총액행' && Number(x.PDF금액||0)>0 && isRetirementName(x.항목 || x.목));
  if(totalRows.length){
    // 같은 목이 여러 번 잡히는 경우 중복 방지를 위해 목명 기준으로 가장 큰 금액만 사용합니다.
    const byName = new Map();
    for(const x of totalRows){
      const key = norm(x.항목 || x.목);
      const prev = byName.get(key);
      if(!prev || Number(x.PDF금액||0) > Number(prev.PDF금액||0)) byName.set(key, x);
    }
    const used = [...byName.values()];
    return {has:true, amount:used.reduce((s,x)=>s+Number(x.PDF금액||0),0), items:used};
  }
  const calcRows = items.filter(x=>x.구분==='산출기초' && Number(x.PDF금액||0)>0 && (isRetirementName(x.목) || isRetirementName(x.항목)));
  return {has:calcRows.length>0, amount:calcRows.reduce((s,x)=>s+Number(x.PDF금액||0),0), items:calcRows};
}
function retirementVerdict(excelSummary, pdfSummary){
  // v38: 퇴직금은 금액 일치 여부를 검토하지 않습니다.
  // 엑셀에 실제 퇴직 적립금액이 있으면 PDF에 퇴직 관련 목/산출항목이 편성되어 있는지만 확인합니다.
  const ea = Number(excelSummary.amount || 0), pa = Number(pdfSummary.amount || 0);
  if(ea > 0 && pa > 0) return `편성 확인(금액 비교 제외 · PDF ${fmt(pa)})`;
  if(ea > 0 && pa <= 0) return `미편성(엑셀 적립금액 있음 · PDF 퇴직 관련 편성 없음)`;
  if(ea <= 0 && pa > 0) return `참고: PDF에는 퇴직 관련 편성이 있으나 엑셀 적립금액은 없음(금액 비교 제외)`;
  return '해당 없음(엑셀 적립금액 없음 · PDF 퇴직 관련 편성 없음)';
}

function buildPrecheck(report){
  const rows=[];
  const issues=[];
  const salaryObj = report.excel?.salary;
  const ex = salaryObj?.summary;
  const pdfItems = report.pdf?.items || [];
  const totals = pdfItems.filter(x=>x.구분==='총액행');
  const calcs = pdfItems.filter(x=>x.구분==='산출기초');
  const totalByName = (name) => totalRowByName(totals, name)?.PDF금액 || 0;
  const amountByCalc = (re) => calcs.filter(x=>re.test(x.항목)).reduce((s,x)=>s+x.PDF금액,0);
  const addIssue = (지적내용, 근거) => issues.push({번호:issues.length+1, 지적내용, 근거});

  // 인건비 산출기초가 '단가*인원*월수' 구조를 갖추지 못한 경우를 지적합니다.
  addBasisIssues(issues, calcs);

  if(ex){
    const teacherBase = ex.교원.소계본봉 || ex.교원.본봉;
    const staffBase = ex.직원.소계본봉 || ex.직원.본봉;

    const teacherPayCalcs = exactMokCalcs(calcs, '교원급여');
    const teacherPayTotal = totalByName('교원급여') || sumAmount(teacherPayCalcs);
    const teacherPayPeople = sumPeopleDistinct(teacherPayCalcs) || '';
    const teacherOff = findOffMokMatches(calcs, '교원급여', amountDiff(teacherBase, teacherPayTotal), '교원급여');
    rows.push({구분:'급여', 항목:'교원급여', 엑셀금액:teacherBase, 엑셀인원:ex.교원.인원, PDF금액:teacherPayTotal, PDF인원:teacherPayPeople, 검토결과:detailVerdict('교원급여', teacherBase, teacherPayTotal, ex.교원.인원, teacherPayPeople, teacherOff)});

    const staffPayCalcs = exactMokCalcs(calcs, '직원급여');
    const staffPayTotal = totalByName('직원급여') || sumAmount(staffPayCalcs);
    const staffPayPeople = sumPeopleDistinct(staffPayCalcs) || '';
    const staffOff = findOffMokMatches(calcs, '직원급여', amountDiff(staffBase, staffPayTotal), '직원급여');
    rows.push({구분:'급여', 항목:'직원급여', 엑셀금액:staffBase, 엑셀인원:ex.직원.인원, PDF금액:staffPayTotal, PDF인원:staffPayPeople, 검토결과:detailVerdict('직원급여', staffBase, staffPayTotal, ex.직원.인원, staffPayPeople, staffOff)});

    // 급여는 세부 직종별로 나누지 않고 교원급여/직원급여 총액과 총인원을 우선 검토합니다.

    const allowanceRows = ex.수당.filter(a=>!LEGAL_RE.test(a.항목));
    const teacherGroup = totalRowByName(totals, '교원수당');
    const staffGroup = totalRowByName(totals, '직원수당');
    const teacherAnalysis = analyzeAllowances('교원', allowanceRows, pdfItems, teacherGroup);
    const staffAnalysis = analyzeAllowances('직원', allowanceRows, pdfItems, staffGroup);

    // 엑셀 수당 항목별 표시: 개별 편성, 통합편성 추정, 미편성/추가확인을 행 단위로 보여줍니다.
    for(const d of teacherAnalysis.details){
      rows.push({구분:'수당', 항목:d.항목, 엑셀금액:d.엑셀금액, 엑셀인원:d.엑셀인원, PDF금액:d.PDF금액, PDF인원:d.PDF인원, 검토결과:d.검토결과});
    }
    for(const d of staffAnalysis.details){
      rows.push({구분:'수당', 항목:d.항목, 엑셀금액:d.엑셀금액, 엑셀인원:d.엑셀인원, PDF금액:d.PDF금액, PDF인원:d.PDF인원, 검토결과:d.검토결과});
    }

    /* v30: 지적사항 표는 퇴직적립금 관련 사항만 표시합니다. 차액/다른 목 편성은 검토결과 열에 표시합니다.
    addOffMokDiffIssues(issues, calcs, [
      {category:'교원급여', expectedMok:'교원급여', excelAmount:teacherBase, pdfAmount:teacherPayTotal},
      {category:'교원수당', expectedMok:'교원수당', excelAmount:teacherAnalysis.totalExcel, pdfAmount:teacherAnalysis.groupMatchedAmount || teacherAnalysis.groupTotal},
      {category:'직원급여', expectedMok:'직원급여', excelAmount:staffBase, pdfAmount:staffPayTotal},
      {category:'직원수당', expectedMok:'직원수당', excelAmount:staffAnalysis.totalExcel, pdfAmount:staffAnalysis.groupMatchedAmount || staffAnalysis.groupTotal}
    ]);

    addMissingAllowanceIssues(issues, teacherAnalysis, staffAnalysis);
    */
  }

  // v30: 일반 오편성/미편성 지적사항은 결과표 검토결과 열에만 표시하고, 지적사항 표에는 퇴직적립금만 표시합니다.

  const excelRetireSummary = getExcelRetirementSummary(report.excel?.retirement || []);
  const pdfRetireSummary = getPdfRetirementSummary(pdfItems);
  rows.push({
    구분:'퇴직',
    항목:'퇴직적립금',
    엑셀금액:excelRetireSummary.amount,
    엑셀인원:'',
    PDF금액:pdfRetireSummary.amount,
    PDF인원:'',
    검토결과:retirementVerdict(excelRetireSummary, pdfRetireSummary)
  });
  if(excelRetireSummary.has && !pdfRetireSummary.has){
    addIssue('엑셀에는 퇴직 적립금액이 있으나 퇴직적립금 미편성', `엑셀 퇴직 관련 시트에서 실제 적립금액이 확인되었으나 PDF에서 퇴직 관련 목/산출항목을 찾지 못했습니다. 퇴직금은 금액 일치 여부는 검토하지 않습니다.`);
  }
  if(!excelRetireSummary.has && pdfRetireSummary.has){
    const names = pdfRetireSummary.items.map(x=>x.항목 || x.목).filter(Boolean).join(', ');
    addIssue('엑셀 적립금액 없음 / PDF 퇴직 관련 편성 확인', `엑셀 퇴직 관련 시트의 실제 적립금액은 확인되지 않았으나 PDF에서 ${names} 편성이 확인되었습니다. 퇴직금은 금액 비교 대상이 아니므로 필요 시 원자료만 확인하세요.`);
  }
  return {rows, issues};
}

function analyzeAllowances(kind, allowanceRows, pdfItems, groupRowOrTotal){
  const amountField = kind === '교원' ? '교원금액' : '직원금액';
  const peopleField = kind === '교원' ? '교원인원' : '직원인원';
  const groupRow = typeof groupRowOrTotal === 'object' ? groupRowOrTotal : {PDF금액:Number(groupRowOrTotal||0)};
  const groupTotal = Number(groupRow?.PDF금액 || 0);
  const relevant = allowanceRows.filter(a=>Number(a[amountField]||0)>0);
  const direct = [];
  const missing = [];
  const details = [];

  for(const a of relevant){
    const amount = Number(a[amountField]||0);
    const matches = findAllowancePdfMatches(kind, a.항목, amount, pdfItems);
    const pdfAmount = matches.reduce((s,x)=>s+Number(x.PDF금액||0),0);
    const pdfPeople = matches.reduce((s,x)=>s+Number(x.인원||0),0);
    if(matches.length){
      direct.push(a);
      details.push({항목:`${kind} ${a.항목}`, 엑셀금액:amount, 엑셀인원:a[peopleField]||'', PDF금액:pdfAmount, PDF인원:pdfPeople||'', 검토결과:`개별편성 확인(${matches.map(m=>m.항목).join(', ')})`});
    }else{
      missing.push(a);
    }
  }

  const totalExcel = relevant.reduce((s,a)=>s+Number(a[amountField]||0),0);
  const missingSum = missing.reduce((s,a)=>s+Number(a[amountField]||0),0);
  const directSum = direct.reduce((s,a)=>s+Number(a[amountField]||0),0);
  const directPdfSum = details
    .filter(d => /^개별편성/.test(d.검토결과 || ''))
    .reduce((s,d)=>s+Number(d.PDF금액||0),0);
  const residual = Math.max(0, groupTotal - directPdfSum);
  const candidates = allowanceGroupCandidates(groupRow, kind);
  if(residual) candidates.unshift({label:`${kind}수당 잔액`, amount:residual});
  let matched = null;
  // 미개별 수당이 있으면 전체 수당 총액이 아니라 "교원/직원수당 총액 - 이미 개별확인된 금액" 또는 미개별 수당 합계와 비교합니다.
  if(missingSum) matched = candidates.find(c=>closeMoney(c.amount, missingSum));
  // 미개별 수당이 없는 경우에만 전체 수당 총액 일치를 인정합니다.
  if(!matched && !missing.length && totalExcel) matched = candidates.find(c=>closeMoney(c.amount,totalExcel));
  const coveredByGroup = !!matched;
  const integratedMessage = coveredByGroup && missing.length ? integratedText(kind, missing, amountField, `${kind}수당`, matched.amount) : '';

  for(const a of missing){
    const amount = Number(a[amountField]||0);
    let msg = '';
    if(integratedMessage) msg = integratedMessage;
    else if(!groupTotal) msg = `추가확인필요: PDF ${kind}수당 총액행 없음`;
    else msg = `추가확인필요: PDF 개별 항목 없음, 미개별표시 수당 합계 ${fmt(missingSum)} / PDF ${kind}수당 ${fmt(groupTotal)}`;
    details.push({항목:`${kind} ${a.항목}`, 엑셀금액:amount, 엑셀인원:a[peopleField]||'', PDF금액:0, PDF인원:0, 검토결과:msg});
  }

  let verdict = '';
  if(integratedMessage) verdict = integratedMessage;
  else if(!missing.length && totalExcel && candidates.some(c=>closeMoney(c.amount,totalExcel))) verdict = `총액 일치 또는 통합편성(${fmt(totalExcel)} ≒ ${kind}수당 ${fmt(groupTotal)})`; 
  else if(missingSum && !groupTotal) verdict = `추가확인필요: PDF ${kind}수당 총액행 없음, 미개별표시 수당 ${fmt(missingSum)}`;
  else if(missingSum) verdict = `추가확인필요: 미개별표시 수당 합계 ${fmt(missingSum)}, PDF ${kind}수당 ${fmt(groupTotal)}`;
  else verdict = '개별 또는 총액 확인';
  return {kind, amountField, peopleField, relevant, direct, missing, details, totalExcel, missingSum, directSum, groupTotal, groupMatchedAmount:matched?.amount||0, coveredByGroup, verdict};
}

function allowancePdfHit(kind, key, amount, x){
  if(!x || !key) return false;
  if(LEGAL_RE.test(x.항목||'')) return false;
  const mok = norm(x.목 || x.상위항목 || '');
  const item = norm(x.항목 || '');
  const calc = norm(x.산출기초 || '');
  const k2 = allowanceKey(x.항목 || '');
  const allowanceMok = kind === '교원' ? '교원수당' : '직원수당';

  // 수당 개별편성은 원칙적으로 해당 수당 목 안에서만 인정합니다.
  // 단, 산출항목 자체가 교원/직원 접두어를 가진 명확한 수당명인 경우도 인정합니다.
  const inRightMok = mok === allowanceMok || mok.includes(allowanceMok);
  const prefixedRightKind = kind === '교원'
    ? /^교원|^교사/.test(item)
    : /^직원|^사무직원|^조리직원|^보조교사|^차량기사|^차량보조|^환경미화|^영양사/.test(item);

  // 식대 ↔ 정액급식비 동의어는 교원/직원 수당 목 또는 명확한 교원/직원 정액급식 항목에서만 적용합니다.
  const isMealKey = key === '정액급식' || k2 === '정액급식';
  if(isMealKey){
    const scopedMeal = inRightMok || item.includes(kind + '정액급식') || item.includes(kind + '식대');
    if(!scopedMeal) return false;
  }

  const keyMatched = !!(k2 && (k2 === key || k2.includes(key) || key.includes(k2)));
  if(!keyMatched) return false;

  // 기타수당처럼 일반적인 명칭은 해당 수당 목 안에 있을 때만 인정합니다.
  if(key.length <= 2 && !inRightMok) return false;

  return inRightMok || prefixedRightKind;
}

function addMissingAllowanceIssues(issues, teacherAnalysis, staffAnalysis){
  const byName = new Map();
  for(const analysis of [teacherAnalysis, staffAnalysis]){
    // 통합수당 총액으로 설명되는 경우는 미편성으로 지적하지 않습니다.
    if(analysis.coveredByGroup) continue;
    for(const a of analysis.missing){
      const name = a.항목;
      if(!byName.has(name)) byName.set(name, {항목:name, 교원금액:0, 직원금액:0});
      byName.get(name)[analysis.kind+'금액'] += a[analysis.amountField] || 0;
    }
  }
  for(const v of byName.values()){
    const parts=[];
    if(v.교원금액) parts.push(`교원 ${fmt(v.교원금액)}`);
    if(v.직원금액) parts.push(`직원 ${fmt(v.직원금액)}`);
    issues.push({번호:issues.length+1, 지적내용:`${v.항목} 예산 미편성`, 근거:`엑셀 보수기준에는 ${parts.join(', ')}이 있으나 PDF 개별 항목 또는 통합수당 총액으로 확인되지 않았습니다.`});
  }
}

function allowanceVerdict(kind, allowanceRows, pdfItems, groupTotal){
  return analyzeAllowances(kind, allowanceRows, pdfItems, groupTotal).verdict;
}
function allowanceKey(s){
  let k = norm(s);
  if(!k) return '';
  // [방], [급], [운영]처럼 기관 내부 구분 태그는 같은 수당명으로 봅니다.
  k = k.replace(/^(?:\[[^\]]+\])+/g, '');
  // 소속/직종 접두어는 같은 항목 매칭에서 제외합니다.
  k = k.replace(/^(교원|직원|교사|사무직원|조리직원|보조교사)/,'');
  // 접두어 제거 뒤에 다시 태그가 오는 경우도 보정합니다. 예: 교원[방]근속수당
  k = k.replace(/^(?:\[[^\]]+\])+/g, '');

  // 기관마다 다른 명칭을 같은 수당으로 봅니다.
  if(/^(식대|급식비|정액급식비|정액급식|급식수당)$/.test(k) || /정액급식/.test(k)) return '정액급식';
  if(/성과상여/.test(k)) return '성과상여';
  if(/^상여금?$/.test(k) || /상여/.test(k)) return '상여';
  if(/명절휴가/.test(k)) return '명절휴가';
  if(/스승의날/.test(k)) return '스승의날상여';
  if(/방학휴가/.test(k)) return '방학휴가';
  if(/자가운전|자가차량|차량유지/.test(k)) return '자가운전';
  if(/직책급|직책/.test(k)) return '직책급';

  const stripped = k.replace(/수당|보조금|지원비|지원금|휴가비|급식비|식대|비/g,'');
  // 전부 지워져 빈 키가 되면 원 명칭을 보수적으로 유지합니다.
  return stripped || k;
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
      if(ex.salaryCandidates?.length){
        html += '<p class="muted">후보 시트에서 직명/본봉/지급액계가 보이지 않으면 엑셀 저장 형식 또는 병합 헤더 문제입니다.</p>';
      }
    }
    html += '<h3 class="section-title">퇴직 관련 시트</h3>'+table(['sheetName','positiveAmount','hasRetirementAmount','retireColumns','positiveCells'], ex.retirement.map(r=>({
      sheetName:r.sheetName,
      positiveAmount:fmt(r.positiveAmount),
      hasRetirementAmount:r.hasRetirementAmount?'있음':'없음',
      retireColumns:(r.retireColumns||[]).map(c=>`${c.열}:${c.헤더}`).join(' / '),
      positiveCells:(r.positiveCells||[]).map(c=>`${c.행}행 ${c.열}열 ${fmt(c.값)}`).join(' / ')
    })));
    $('excelTables').innerHTML = html;
  }
  if(report.pdf){
    const pdf=report.pdf;
    $('pdfSummary').innerHTML = `<span class="pill">파일 ${escapeHtml(pdf.fileName)}</span><span class="pill">${pdf.pageCount}페이지</span><span class="pill">항목 ${pdf.items.length}개</span>`;
    let html = '<h3 class="section-title">PDF 추출 항목</h3>'+table(['구분','페이지','목','항목','PDF금액','PDF금액천원','누락항목','산출기초'], pdf.items.map(x=>({...x,PDF금액:fmt(x.PDF금액),누락항목:(x.누락항목||[]).join(', ')})));
    html += '<h3 class="section-title">PDF 원문 라인</h3>'+table(['페이지','y','텍스트'], pdf.lines.slice(0,300), {small:true});
    $('pdfTables').innerHTML = html;
  }
  const review = report.precheck || {rows:[], issues:[]};
  let reviewHtml = '<h3 class="section-title">금액 검토</h3>' + table(['구분','항목','엑셀금액','PDF금액','검토결과'], review.rows.map(r=>({구분:r.구분, 항목:r.항목, 엑셀금액:fmt(r.엑셀금액), PDF금액:fmt(r.PDF금액), 검토결과:r.검토결과})));
  reviewHtml += '<h3 class="section-title">지적사항</h3>' + table(['번호','지적내용','근거'], review.issues || []);
  $('precheck').innerHTML = reviewHtml;
}
