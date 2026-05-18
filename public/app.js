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
  let items = extractBudgetItems(lines);
  items = enrichBudgetItemsFromSplitLines(lines, items);
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

    if(!isCalcStartCandidate(t)) continue;

    const block = collectCalcBlock(lines, i);
    if(block && block.calc){
      let name = block.calc.name || '';
      const nearName = findCalcName(lines, i) || findCalcName(lines, block.end) || '';
      if(!name || name === currentTotal?.항목 || name === '산출항목미상' || name.length < 2 || /^[가-힣]$/.test(name)) name = nearName;
      if(!name) name = currentTotal?.항목 || '산출항목미상';
      out.push({페이지:l.page, 행:i+1, 목:currentTotal?.항목 || '', 상위항목:currentTotal?.항목 || '', 항목:name, PDF금액:block.calc.amount, PDF금액천원:Math.round(block.calc.amount/1000), 인원:block.calc.people, 누락항목:block.calc.missing || [], 산출기초:block.basis, 구분:'산출기초'});
      i = block.end;
    }
  }
  return out.filter(x=>!LEGAL_RE.test(x.항목));
}

function budgetItemId(x){
  return [x.구분||'', x.페이지||'', norm(x.목||x.상위항목||''), norm(cleanItemName(x.항목||'')), Math.round(Number(x.PDF금액||0))].join('|');
}
function likelySplitNameLine(s){
  const t = text(s||'');
  if(!t || /\(본예산\)|=|\*/.test(t)) return false;
  if(parseTotalBudgetRow(t)) return false;
  if(/발행일|예산구분|과\s*목|산출내역|산출기초|보조금|수익자|합계/.test(t)) return false;
  return /퇴직|급보조비|구수당|장근로수당|급식비|통보조비|절휴가비|근속수당|상여금|기타수당|직적립금|본급|립금/.test(t.replace(/\s+/g,''));
}
function recoverSplitTotalRows(lines, existing){
  const out=[];
  const seen = new Set((existing||[]).map(budgetItemId));
  for(let i=0;i<lines.length-2;i++){
    const a = text(lines[i].text||''), b = text(lines[i+1].text||''), c = text(lines[i+2].text||'');
    if(lines[i].page !== lines[i+1].page || lines[i].page !== lines[i+2].page) continue;
    const compactName = norm(a + c);
    if(!/퇴직/.test(compactName)) continue;
    const nums = [...b.matchAll(/-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+/g)].map(m=>toNum(m[0]));
    // 세출명세서 총액행의 재원별 숫자 4개를 우선 사용합니다.
    if(nums.length < 4) continue;
    const name = normalizeMokName(a + c);
    if(!validBudgetRowName(name)) continue;
    const item = {페이지:lines[i].page, 행:i+1, 목:name, 상위항목:'', 항목:name, PDF금액:nums[3]*1000, PDF금액천원:nums[3], 보조금금액:nums[0]*1000, 수익자금액:nums[1]*1000, 기타금액:nums[2]*1000, 인원:null, 산출기초:'', 구분:'총액행', 보정:'split-total'};
    const id = budgetItemId(item);
    if(!seen.has(id)){ seen.add(id); out.push(item); }
  }
  return out;
}
function recoverSplitCalcRows(lines, existing){
  const out=[];
  const seen = new Set((existing||[]).map(budgetItemId));
  let currentTotal = null;
  for(let i=0;i<lines.length;i++){
    const t = text(lines[i].text||'');
    const total = parseTotalBudgetRow(t);
    if(total) currentTotal = {항목:total.항목, 페이지:lines[i].page, 행:i+1, 금액:total.금액};
    // split total row such as 교원퇴직금및퇴직적 / number row / 립금
    if(i+2<lines.length && lines[i].page===lines[i+2].page){
      const combined = norm((lines[i].text||'') + (lines[i+2].text||''));
      if(/퇴직/.test(combined)){
        const nums = [...String(lines[i+1].text||'').matchAll(/-?[0-9]{1,3}(?:,[0-9]{3})*|-?[0-9]+/g)].map(m=>toNum(m[0]));
        if(nums.length>=4) currentTotal = {항목:normalizeMokName(String(lines[i].text||'')+String(lines[i+2].text||'')), 페이지:lines[i].page, 행:i+1, 금액:nums[3]*1000};
      }
    }
    if(!/\(본예산\)|[0-9,]+\s*원?\s*\*/.test(t)) continue;
    if(!/=\s*$|=\s*[^0-9]*$|[0-9,]+\s*원?\s*\*/.test(t)) continue;
    const buf=[t];
    let end=i;
    for(let k=i+1;k<Math.min(lines.length,i+8);k++){
      if(lines[k].page !== lines[i].page) break;
      const s = text(lines[k].text||'');
      if(!s) continue;
      if(k>i && parseTotalBudgetRow(s)) break;
      if(k>i && /\(본예산\)/.test(s)) break;
      buf.push(s); end=k;
      const joined = buf.join(' ').replace(/\s+/g,' ').trim();
      const calc = parseCalcLine(joined);
      if(calc && Number(calc.amount||0)>0){
        let name = calc.name || knownCalcNameFromBasis(joined) || findCalcName(lines, i) || findCalcName(lines, end) || currentTotal?.항목 || '산출항목미상';
        // 줄바꿈으로 첫 글자가 앞줄에 붙고 나머지가 다음 줄로 넘어간 항목 보정
        const known = knownCalcNameFromBasis(joined);
        if(known) name = known;
        name = cleanItemName(name);
        const item = {페이지:lines[i].page, 행:i+1, 목:currentTotal?.항목 || '', 상위항목:currentTotal?.항목 || '', 항목:name, PDF금액:calc.amount, PDF금액천원:Math.round(calc.amount/1000), 인원:calc.people, 누락항목:calc.missing || [], 산출기초:joined, 구분:'산출기초', 보정:'split-calc'};
        const id = budgetItemId(item);
        if(!seen.has(id)){ seen.add(id); out.push(item); }
        break;
      }
      // 금액 줄까지 왔는데 항목명만 있고 산식이 아니면 계속 보지 않습니다.
      if(k>i && /[0-9]{1,3}(?:,[0-9]{3})+\s*$/.test(s) && likelySplitNameLine(s)) break;
    }
  }
  return out;
}
function enrichBudgetItemsFromSplitLines(lines, items){
  const base = Array.isArray(items) ? items.slice() : [];
  const additions = [...recoverSplitTotalRows(lines, base), ...recoverSplitCalcRows(lines, base)];
  const seen = new Set(base.map(budgetItemId));
  for(const x of additions){
    const id = budgetItemId(x);
    if(!seen.has(id)){ seen.add(id); base.push(x); }
  }
  return base.filter(x=>!LEGAL_RE.test(x.항목));
}

function isCalcStartCandidate(t){
  const s = text(t || '');
  if(!s) return false;
  if(/예산구분|발행일|과\s*목|산출내역|산출기초|보조금\s*및|수익자\s*부담/.test(s)) return false;
  if(parseTotalBudgetRow(s)) return false;
  // 산출내역은 보통 (본예산)으로 시작하지만, PDF 추출상 항목명과 산식만 남는 경우도 있어 산식 줄도 후보로 둡니다.
  return /\(본예산\)|=\s*[0-9,]+\s*$|[0-9,]+\s*원?\s*\*/.test(s);
}
function collectCalcBlock(lines, startIdx){
  const buf=[];
  const startPage = lines[startIdx]?.page;
  const maxEnd = Math.min(lines.length-1, startIdx + 22);
  for(let k=startIdx;k<=maxEnd;k++){
    const lt = lines[k];
    if(!lt || lt.page !== startPage) break;
    const s = text(lt.text || '');
    if(!s) continue;
    if(k>startIdx && parseTotalBudgetRow(s)) break;
    if(k>startIdx && /\(본예산\)/.test(s) && parseCalcLine(buf.join(' '))){
      break;
    }
    // 새 산출내역이 시작됐는데 현재 블록이 아직 금액으로 끝나지 않으면 과도하게 합치지 않습니다.
    if(k>startIdx && /\(본예산\)/.test(s) && buf.some(x=>/=\s*[0-9,]+\s*$/.test(x))){
      break;
    }
    buf.push(s);
    const joined = buf.join(' ').replace(/\s+/g,' ').trim();
    const calc = parseCalcLine(joined);
    if(calc){
      return {start:startIdx, end:k, calc, basis:joined};
    }
  }
  // 시작 줄 자체가 계산식인 경우를 마지막으로 한 번 더 확인합니다.
  const joined = buf.join(' ').replace(/\s+/g,' ').trim();
  const calc = parseCalcLine(joined);
  return calc ? {start:startIdx, end:startIdx, calc, basis:joined} : null;
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

function knownCalcNameFromBasis(t){
  const raw = text(t || '');
  const compact = norm(raw);
  const patterns = [
    ['[방]근속수당', /\[방\]\s*근속수당|방근속수당/],
    ['근속수당', /(?<!방)근속수당|(^|[^가-힣])근속수당/],
    ['[방]급식수당', /\[방\]\s*급식수당|방급식수당/],
    ['급식수당', /(?<!방)급식수당|정액\s*급식비|급식비\(교원\)|교원정액급식비|직원정액급식비/],
    ['직책수당', /직책수당|직책급업무추진비/],
    // PDF가 "직 150,000원... 급보조비"처럼 산출내역 첫 글자를 산식 앞에 떼어내는 경우 보정
    ['직급보조비', /직급보조비|직급수당|급보조비/],
    ['연구수당', /연구수당|연구활동비|연구비|구수당\(교원\)|구수당/],
    ['연장수당', /연장수당|연장근로수당|시간외수당|장근로수당/],
    ['운전수당', /운전수당|자가운전보조금|자가운전|교통보조|통보조비/],
    ['기타수당', /기타수당/],
    ['성과상여금', /성과상여금/],
    ['스승의날상여금', /스승의날상여금/],
    ['방학휴가비', /방학휴가비/],
    ['명절휴가비', /명절휴가비|절휴가비/],
    ['상여금', /(^|[^가-힣])(상여금)([^가-힣]|$)|\s상여금|상여금/],
    ['퇴직적립금', /퇴직적립금|직적립금|퇴직금및퇴직적립|퇴직금/],
    ['기본급(영양사)', /기본급\s*\(?영양사\)?|영양사급여/],
    ['차량기사급여', /차량기사급여|기사급여/],
  ];
  // 산출기초 블록 안에 여러 항목명이 섞여 있으면, 실제 산식 금액 앞에 가장 가까운 항목명을 우선합니다.
  const eqIdx = raw.lastIndexOf('=');
  const searchArea = eqIdx >= 0 ? raw.slice(0, eqIdx) : raw;
  let best = null;
  for(const [name, re] of patterns){
    const ms = [...searchArea.matchAll(new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g'))];
    if(ms.length){
      const pos = ms[ms.length-1].index ?? 0;
      if(!best || pos > best.pos) best = {name, pos};
    }
  }
  if(best) return best.name;
  if(/상여금/.test(compact)) return '상여금';
  return '';
}

function extractCalcItemNameFromLine(t){
  let s = text(t || '');
  s = s.replace(/\(본예산\)|\(보조금및지원금\)|\(수익자부담금\)|\(그밖의수입\)/g, ' ')
       .replace(/\s+/g, ' ')
       .trim();
  if(!s) return '';
  // 산출내역 여러 줄 + 산출기초 여러 줄이 합쳐진 블록에서 산식 시작 전까지를 항목명으로 봅니다.
  // 예: "직책수당 (원장500,000원+부원장200,000원+교사50,000원)*12월 = 9,000,000"
  // 예: "상여금 {(명절 218,182*11명)+...}=12,901,000"
  let m = s.match(/^(.{1,80}?)(?=\s*(?:[\{\(][^=]{0,80}?[0-9,]+\s*원?|[0-9,]+\s*원?\s*\*))/);
  if(!m) return '';
  let name = cleanCalcNameCandidate(m[1]);
  const known = knownCalcNameFromBasis(t);
  if(known && (!name || name.length > 24 || /명절|행사|방학|스승/.test(name))) return known;
  if(!name) return known || '';
  return name;
}
function parseCalcLine(t){
  let extractedName = extractCalcItemNameFromLine(t);
  const knownName = knownCalcNameFromBasis(t);
  if(knownName && (!extractedName || /명절|행사|방학|스승/.test(extractedName))) extractedName = knownName;
  const compact = text(t || '').replace(/\s+/g,'').replace(/＝/g,'=').replace(/[{}]/g,'');
  let m = compact.match(/([0-9,]+)원\*([0-9,]+)명\*([0-9,]+)(?:월|개월|개?월)=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:toNum(m[2]), months:toNum(m[3]), amount:toNum(m[4]), missing:[], name:extractedName};

  // 인원 수가 빠진 형태: 3,600,000원*12월=43,200,000
  m = compact.match(/([0-9,]+)원\*([0-9,]+)(?:월|개월|개?월)=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:null, months:toNum(m[2]), amount:toNum(m[3]), missing:['인원 수'], name:extractedName};

  // 월수/개월 수가 빠진 형태: 100,000원*5명=500,000
  m = compact.match(/([0-9,]+)원\*([0-9,]+)명=([0-9,]+)$/);
  if(m) return {unit:toNum(m[1]), people:toNum(m[2]), months:null, amount:toNum(m[3]), missing:['월수'], name:extractedName};

  // 복합 괄호 산식: (원장500,000원+부원장200,000원+교사50,000원)*12월=9,000,000
  // 또는 원 단위 표기가 일부 빠진 상여금 산식도 금액과 항목명을 우선 잡습니다.
  // 생각키움 PDF처럼 '= 급보조비(교원) 7,200,000' 형태로 산출내역명이 등호 뒤에 끼는 경우도 허용합니다.
  if(/=[^0-9]{0,80}([0-9,]+)(?:원)?$/.test(compact)){
    const amountMatch = compact.match(/=[^0-9]{0,80}([0-9,]+)(?:원)?$/);
    const peopleMatches = [...compact.matchAll(/([0-9,]+)명/g)].map(m=>toNum(m[1])).filter(Boolean);
    const monthMatch = compact.match(/([0-9,]+)(?:월|개월|개?월)/);
    const unitMatch = compact.match(/([0-9,]+)원/);
    const missing = [];
    if(!peopleMatches.length) missing.push('인원 수');
    if(!monthMatch && !/([0-9,]+)회/.test(compact)) missing.push('월수');
    return {
      unit: unitMatch ? toNum(unitMatch[1]) : null,
      people: peopleMatches.length ? Math.max(...peopleMatches) : null,
      months: monthMatch ? toNum(monthMatch[1]) : null,
      amount: toNum(amountMatch[1]),
      missing,
      name: extractedName
    };
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
  // PDF 텍스트 추출에서는 산출항목명이 산출기초의 바로 위에 오기도 하고,
  // 바로 아래에 오기도 합니다. 한쪽 방향만 보면 직원방학휴가비처럼
  // 다른 항목명과 잘못 매칭되는 문제가 있어 가까운 양방향 후보를 봅니다.
  const offsets = [1,-1,2,-2,3,-3,4,-4,5,-5,6,-6];
  for(const off of offsets){
    const j = idx + off;
    if(j < 0 || j >= lines.length) continue;
    const cand = cleanCalcNameCandidate(lines[j]?.text || '');
    if(cand) return cand;
  }
  return '';
}


function offMokMatchScore(x, category, expectedMok){
  const label = norm(`${cleanItemName(x?.항목 || '')} ${x?.산출기초 || ''} ${x?.목 || x?.상위항목 || ''}`);
  const cat = norm(category || expectedMok || '');
  let score = 0;
  if(/직원급여|급여/.test(cat)){
    if(/차량기사급여|기사급여|운전.*급여|급여.*기사/.test(label)) score += 100;
    if(/영양사|기본급영양사|영양사급여/.test(label)) score += 100;
    if(/급여|인건비|보수/.test(label)) score += 35;
    if(/주유|유류|운영비|수리|보험|임차료/.test(label)) score -= 25;
  }
  if(/교원급여/.test(cat) && /교원급여|교사급여|원장급여|방과후교원급여/.test(label)) score += 80;
  if(/수당/.test(cat) && /(수당|상여|휴가|급식|자가운전|직급|직책|연구|관리업무)/.test(label)) score += 50;
  // 항목명에 category의 의미어가 직접 포함되면 보강
  for(const token of ['차량기사','보조교사','조리직원','사무직원','환경미화','영양사','교원','직원']){
    if(cat.includes(norm(token)) && label.includes(norm(token))) score += 30;
  }
  return score;
}


function findOffMokMatches(calcs, expectedMok, diff, category){
  const absDiff = Math.abs(Number(diff||0));
  if(absDiff <= 1000) return [];
  const target = norm(expectedMok);
  const categoryNorm = norm(category || expectedMok);
  const isPay = /급여/.test(category);
  const isAllowance = /수당/.test(category);
  const direct = calcs.filter(x=>{
    if(!x || !x.PDF금액) return false;
    if(LEGAL_RE.test(x.항목 || '')) return false;
    const mok = norm(x.목 || x.상위항목 || '');
    if(!mok || mok.includes(target)) return false;
    if(!closeMoney(Math.abs(x.PDF금액), absDiff)) return false;
    const label = norm((x.항목 || '') + ' ' + (x.산출기초 || ''));
    if(isPay){
      // 직원급여 차액이 다른 목의 기본급/급여/인건비성 산출내역에 숨어 있는 경우를 잡습니다.
      // 예: 그밖의인건비 > 기본급(영양사), 통학차량이용비 > 차량기사급여
      if(/직원급여/.test(categoryNorm) && /(통학차량|차량|임차료|운행|기사|영양사|기본급|급여|인건비)/.test(label + mok)) return true;
      return /(급여|기본급|인건비|보수|영양사)/.test(label);
    }
    if(isAllowance) return /(수당|상여|휴가비|식대|정액급식|자가운전|직급|직책|연구|관리업무)/.test(label);
    return true;
  });
  if(direct.length){
    // 같은 금액의 산출내역이 여러 개 있으면 금액만 보지 않고 명칭 유사도를 우선합니다.
    // 예: 18,000,000원이 차량주유비와 차량기사급여에 모두 있을 때, 직원급여 차액은 차량기사급여를 우선 선택.
    direct.sort((a,b)=>offMokMatchScore(b, category, expectedMok) - offMokMatchScore(a, category, expectedMok));
    return direct;
  }

  // v49: 차액이 한 산출항목이 아니라 같은 목 안의 여러 산출항목 합계와 일치하는 경우를 잡습니다.
  // 예: 직원급여 차액 69,300,000원 = 통학차량이용비 목의 통학차량임차료(수익자) 2,004,000원 + 통학차량임차료 67,296,000원
  if(/직원급여|교원급여/.test(categoryNorm)){
    const groups = new Map();
    for(const x of calcs){
      if(!x || !x.PDF금액 || LEGAL_RE.test(x.항목 || '')) continue;
      const mokText = String(x.목 || x.상위항목 || '');
      const mok = norm(mokText);
      if(!mok || mok.includes(target)) continue;
      const label = norm((x.항목 || '') + ' ' + (x.산출기초 || '') + ' ' + mokText);
      if(!/(통학차량|차량|임차료|운행|기사)/.test(label)) continue;
      if(!groups.has(mok)) groups.set(mok, {목:mokText, amount:0, items:[]});
      const g = groups.get(mok);
      g.amount += Number(x.PDF금액 || 0);
      g.items.push(x);
    }
    for(const g of groups.values()){
      if(closeMoney(g.amount, absDiff)){
        const names = [...new Set(g.items.map(x=>cleanItemName(x.항목)).filter(Boolean))];
        const mainName = names.find(n=>/통학|차량|임차/.test(norm(n))) || names[0] || '통학차량임차료';
        return [{
          페이지:g.items[0]?.페이지 || '', 행:g.items[0]?.행 || '', 목:g.목,
          항목:mainName, PDF금액:g.amount, 산출기초:g.items.map(x=>x.산출기초).join(' / '), 구분:'산출기초합계'
        }];
      }
    }
  }
  return [];
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
  // pdf.js가 '차량...'의 첫 글자 '차'를 앞 블록으로 떼어내는 경우 보정
  if(/^량기사/.test(s)) s = '차' + s;
  if(/^량보조/.test(s)) s = '차' + s;
  if(/^량주유/.test(s)) s = '차' + s;
  if(/^량운영/.test(s)) s = '차' + s;
  if(/^량임차/.test(s)) s = '차' + s;
  if(/^직적립금/.test(s)) s = '퇴' + s;
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
      const itemName = cleanItemName(m.항목);
      const offText = (/직원급여/.test(label) && /통학차량|차량/.test(norm((m.목||'') + itemName)))
        ? `${m.목 || m.상위항목} 목의 ${itemName} ${fmt(m.PDF금액)}로 편성`
        : `${m.목 || m.상위항목} 목에 ${itemName} ${fmt(m.PDF금액)} 편성`;
      parts.push(`금액 차이(${label} ${shortWon(Math.abs(diff))} 차이 → ${offText})`);
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

function allowanceIntegratedCalcCandidates(kind, pdfItems){
  const allowanceMok = kind === '교원' ? '교원수당' : '직원수당';
  const out=[];
  for(const x of (pdfItems || [])){
    if(x.구분 !== '산출기초') continue;
    const mok = norm(x.목 || x.상위항목 || '');
    const item = norm(cleanItemName(x.항목 || ''));
    if(!(mok === allowanceMok || mok.includes(allowanceMok))) continue;
    // '교원수당', '직원수당'처럼 세부명이 없는 통합 산출내역만 후보로 사용합니다.
    // 직원정액급식비, 직원명절휴가비 등 개별 산출내역은 이미 direct에서 처리되어야 합니다.
    if(item === allowanceMok || item === norm(kind + '수당') || allowanceKey(item) === '수당'){
      out.push({label:cleanItemName(x.항목 || `${kind}수당`), amount:Number(x.PDF금액||0), source:x});
    }
  }
  return out.filter((x,i,a)=>x.amount && a.findIndex(y=>y.amount===x.amount && y.label===x.label)===i);
}

function allowanceKeys(s){
  const raw = String(s || '');
  const compact = norm(raw);
  const keys = new Set();
  // 복합 항목: "연구수당/교통보조금"처럼 엑셀 한 열에 두 성격이 함께 있는 경우
  // PDF에서는 각각 연구수당, 교통보조비로 나뉘어 편성될 수 있으므로 구성요소를 모두 찾습니다.
  raw.split(/[\/,+·ㆍ&]+|및|와|과/).map(x=>x.trim()).filter(Boolean).forEach(part=>{
    const k = allowanceKey(part);
    if(k) keys.add(k);
  });
  const k = allowanceKey(raw);
  if(k) keys.add(k);
  // 문자열 안에 대표 수당명이 같이 들어 있으면 보조 키를 추가합니다.
  if(/연구/.test(compact)) keys.add('연구');
  if(/교통보조|자가운전|운전수당|차량유지/.test(compact)) keys.add('자가운전');
  if(/연장근로|시간외|연장수당/.test(compact)) keys.add('시간외');
  if(/급식|식대|정액급식/.test(compact)) keys.add('정액급식');
  if(/직급보조|직급수당/.test(compact)) keys.add('직급');
  if(/직책/.test(compact)) keys.add('직책급');
  if(/근속/.test(compact)) keys.add('근속');
  // '상여'는 너무 포괄적이어서 스승의날상여금과 성과상여금을 혼동시키므로 보조 키로 추가하지 않습니다.
  return [...keys].filter(Boolean);
}
function findAllowancePdfMatches(kind, allowanceName, amount, pdfItems){
  const keys = allowanceKeys(allowanceName);
  const seen = new Set();
  const matches = [];
  for(const x of pdfItems){
    if(x.구분 !== '산출기초') continue;
    if(!keys.some(key => allowancePdfHit(kind, key, amount, x))) continue;
    const pdfAmt = Number(x.PDF금액 || 0);
    const excelAmt = Number(amount || 0);
    // 같은 수당이 [방]/[급]/재원별 등으로 나뉜 경우에는 여러 블록을 합산해야 하므로
    // 개별 블록은 엑셀 금액 이하이면 허용합니다.
    if(excelAmt && !(closeMoney(pdfAmt, excelAmt) || pdfAmt <= excelAmt + 1000)) continue;
    const id = `${x.페이지}|${x.행}|${x.목}|${x.항목}|${x.PDF금액}`;
    if(seen.has(id)) continue;
    seen.add(id); matches.push(x);
  }

  // 보정: [방]근속수당처럼 같은 목 안에 같은 핵심수당이 분할 편성된 경우,
  // 첫 매칭 후 종료하지 않고 같은 key의 산출내역을 추가로 합산합니다.
  // 단, PDF 전체 탐색으로 번지지 않도록 교원수당/직원수당 목 내부만 허용합니다.
  const allowanceMok = kind === '교원' ? '교원수당' : '직원수당';
  const currentSum = matches.reduce((s,x)=>s+Number(x.PDF금액||0),0);
  if(keys.length && amount && currentSum && !closeMoney(currentSum, amount)){
    for(const x of pdfItems){
      if(x.구분 !== '산출기초') continue;
      const mok = norm(x.목 || x.상위항목 || '');
      if(!(mok === allowanceMok || mok.includes(allowanceMok))) continue;
      const id = `${x.페이지}|${x.행}|${x.목}|${x.항목}|${x.PDF금액}`;
      if(seen.has(id)) continue;
      const xKeys = allowanceKeys(`${x.항목 || ''} ${x.산출기초 || ''}`);
      if(!keys.some(k => xKeys.some(xk => sameAllowanceKeyForMatch(k, xk)))) continue;
      const projected = matches.reduce((s,m)=>s+Number(m.PDF금액||0),0) + Number(x.PDF금액||0);
      if(projected <= Number(amount||0) + 1000 || closeMoney(projected, amount)){
        seen.add(id); matches.push(x);
        if(closeMoney(projected, amount)) break;
      }
    }
  }
  return matches;
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
  return /(퇴직|직적립)/.test(n) && /(적립|직적립|퇴직금|충당|급여|퇴직)/.test(n);
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
      const names = [...new Set(matches.map(m=>cleanItemName(m.항목)).filter(Boolean))].join(', ');
      const diff = Math.abs(Math.round(amount - pdfAmount));
      const result = closeMoney(amount, pdfAmount)
        ? `개별편성확인(금액일치: ${names})`
        : `개별편성확인(${fmt(diff)} 차이: ${names})`;
      details.push({항목:`${kind} ${a.항목}`, 엑셀금액:amount, 엑셀인원:a[peopleField]||'', PDF금액:pdfAmount, PDF인원:pdfPeople||'', 검토결과:result});
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
  // 통합편성 비교는 목 전체 총액이 아니라 산출내역 중 '교원수당/직원수당' 블록 또는
  // 총액에서 이미 개별확인된 금액을 뺀 잔액만 사용합니다.
  // 예: 직원수당 목 총액 37,600,000원을 통합수당으로 보지 않고, 산출내역 '직원수당' 6,600,000원만 사용.
  const candidates = allowanceIntegratedCalcCandidates(kind, pdfItems);
  if(residual) candidates.unshift({label:`${kind}수당 통합편성금액`, amount:residual});
  // 산출내역 후보가 전혀 없는 예외 파일에서만 재원별 금액을 보조 후보로 사용하되, 목 전체 총액은 제외합니다.
  if(!candidates.length){
    for(const c of allowanceGroupCandidates(groupRow, kind).filter(c=>!closeMoney(c.amount, groupTotal))){
      candidates.push(c);
    }
  }
  let matched = null;
  // 미개별 수당이 있으면 전체 수당 총액이 아니라 "교원/직원수당 총액 - 이미 개별확인된 금액" 또는 미개별 수당 합계와 비교합니다.
  if(missingSum) matched = candidates.find(c=>closeMoney(c.amount, missingSum));
  // 미개별 수당이 없는 경우에만 전체 수당 총액 일치를 인정합니다.
  if(!matched && !missing.length && totalExcel) matched = candidates.find(c=>closeMoney(c.amount,totalExcel));
  const coveredByGroup = !!matched;
  const integratedMessage = coveredByGroup && missing.length ? integratedText(kind, missing, amountField, matched.label || `${kind}수당`, matched.amount) : '';

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


function sameAllowanceKeyForMatch(a,b){
  if(!a || !b) return false;
  if(a === b) return true;
  // 아래 핵심 키워드는 서로 섞이면 안 됩니다.
  const strict = ['스승의날상여','성과상여','명절휴가','방학휴가','정액급식','시간외','직급','직책급','자가운전','연구','관리업무'];
  if(strict.includes(a) || strict.includes(b)) return a === b;
  // '수당'은 통합 산출내역 후보에서만 쓰고 개별 수당 매칭에는 쓰지 않습니다.
  if(a === '수당' || b === '수당') return false;
  return a.length >= 3 && b.length >= 3 && (a.includes(b) || b.includes(a));
}

function allowancePdfHit(kind, key, amount, x){
  if(!x || !key) return false;
  if(LEGAL_RE.test(x.항목||'')) return false;
  const mok = norm(x.목 || x.상위항목 || '');
  const item = norm(x.항목 || '');
  const labelText = `${x.항목 || ''} ${x.산출기초 || ''}`;
  const k2 = allowanceKey(x.항목 || '');
  const labelKeys = allowanceKeys(labelText);
  const allowanceMok = kind === '교원' ? '교원수당' : '직원수당';

  // 수당 개별편성은 원칙적으로 해당 수당 목 안에서만 인정합니다.
  // 단, 산출항목 자체가 교원/직원 접두어를 가진 명확한 수당명인 경우도 인정합니다.
  const inRightMok = mok === allowanceMok || mok.includes(allowanceMok);
  const prefixedRightKind = kind === '교원'
    ? /^교원|^교사/.test(item)
    : /^직원|^사무직원|^조리직원|^보조교사|^차량기사|^차량보조|^환경미화|^영양사/.test(item);

  // 식대 ↔ 정액급식비 동의어는 교원/직원 수당 목 또는 명확한 교원/직원 정액급식 항목에서만 적용합니다.
  const isMealKey = key === '정액급식' || k2 === '정액급식' || labelKeys.includes('정액급식');
  if(isMealKey){
    const scopedMeal = inRightMok || item.includes(kind + '정액급식') || item.includes(kind + '식대');
    if(!scopedMeal) return false;
  }

  // v49: PDF 산출내역이 여러 줄로 쪼개지는 파일에서는 x.항목만 보면 '직', '급보조비'처럼 일부가 빠질 수 있습니다.
  // 산출기초 블록 전체(labelText)에서 동의어 키를 다시 추출해 같은 수당 여부를 판단합니다.
  const keyMatched = labelKeys.some(k => sameAllowanceKeyForMatch(key, k));
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
  // norm()이 대괄호를 제거하므로 [방]근속수당은 "방근속수당"처럼 됩니다.
  // 기관 내부 접두어는 같은 수당명으로 보되, 방학휴가비 같은 실제 단어는 보존합니다.
  k = k.replace(/^(교원|직원|교사|사무직원|조리직원|보조교사)/,'');
  k = k.replace(/^(방|급|운영|수|교)(?=(근속|급식|기본급|상여|연장|연구|직책|직급|운전|기타|식대|정액급식))/,'');

  // 기관마다 다른 명칭을 같은 수당으로 봅니다.
  if(/식대|급식비|정액급식|급식수당|급식보조/.test(k)) return '정액급식';
  if(/연장근로|시간외|연장수당/.test(k)) return '시간외';
  if(/연구활동|연구수당|연구비|연구/.test(k)) return '연구';
  if(/교통보조|교통비|자가운전|자가차량|차량유지|운전수당/.test(k)) return '자가운전';
  if(/직급보조|직급수당/.test(k)) return '직급';
  if(/직책급|직책/.test(k)) return '직책급';
  if(/근속/.test(k)) return '근속';
  if(/성과상여/.test(k)) return '성과상여';
  if(/명절휴가/.test(k)) return '명절휴가';
  if(/스승의날/.test(k)) return '스승의날상여';
  if(/방학휴가/.test(k)) return '방학휴가';
  if(/^상여금?$/.test(k) || k === '상여') return '상여';

  const stripped = k.replace(/수당|보조금|지원비|지원금|상여금|휴가비|급식비|식대|비/g,'');
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

/* =========================
   v50 보강 패치
   - 작성서식 PDF 1차 지원
   - 시트명/내용 기반 탐색 강화
   - 퇴직금 교원/직원 구분 보조
   - 수식 셀 일부 계산 보정
   - 오편성/미편성 지적사항 문구 개선
   ========================= */

const SOURCE_PDF_NOTICE = '작성서식 PDF는 엑셀에서 변환된 텍스트형 PDF만 지원합니다. 스캔본 PDF는 브라우저 정적 방식에서 정확히 읽기 어렵습니다.';

async function parseExcel(file){
  const lower = String(file.name || '').toLowerCase();
  if(lower.endsWith('.pdf')) return await parseSourceFormPdf(file);
  const originalBuf = await file.arrayBuffer();
  const attempts = [];
  try{ attempts.push(await parseExcelWithSheetJS_v50(file.name, originalBuf)); }
  catch(e){ attempts.push({parser:'sheetjs-v50', error:e.message, salaryCandidates:[], salary:null, visibleSheets:[]}); console.warn(e); }
  try{ attempts.push(await parseExcelWithRawXml_v50(file.name, originalBuf)); }
  catch(e){ attempts.push({parser:'raw-xlsx-xml-v50', error:e.message, salaryCandidates:[], salary:null, visibleSheets:[]}); console.warn(e); }
  const successful = attempts.filter(x=>x && x.salary && x.salary.ok);
  let chosen = successful[0];
  if(successful.length > 1) chosen = successful.sort((a,b)=>selectedSalaryScore(b)-selectedSalaryScore(a))[0];
  if(!chosen) chosen = attempts.sort((a,b)=>(b.salaryCandidates?.length||0)-(a.salaryCandidates?.length||0))[0] || {fileName:file.name};
  chosen.fileName = file.name;
  chosen.parserAttempts = attempts.map(a=>({
    parser:a.parser, error:a.error||'', visibleSheets:a.visibleSheets||[],
    salaryFound:!!(a.salary&&a.salary.ok), salarySheet:a.salary?.sheetName||'', candidates:a.salaryCandidates||[]
  }));
  return chosen;
}

async function parseExcelWithSheetJS_v50(fileName, originalBuf){
  const normalizedBuf = await normalizeXlsxForSheetJS(originalBuf);
  const wb = XLSX.read(normalizedBuf, {type:'array', cellDates:false, cellNF:false, cellText:true, raw:false, WTF:false, dense:false});
  const sheetMeta = (wb.Workbook && wb.Workbook.Sheets) || [];
  const visible = wb.SheetNames.map((name,i)=>({name, hidden: sheetMeta[i]?.Hidden || 0})).filter(s=>!s.hidden);
  const sheetAoas = new Map();
  const candidates = visible.map(s=>{
    const aoa = sheetToAoaRobust_v50(wb.Sheets[s.name]);
    sheetAoas.set(s.name, aoa);
    return {...s, score: sheetScore_v50(s.name, aoa)};
  }).filter(s=>s.score>0).sort((a,b)=>b.score-a.score);
  let salary=null; const tried=[];
  for(const c of candidates){
    const parsed = parseSalarySheet(c.name, sheetAoas.get(c.name) || []);
    tried.push({sheet:c.name, score:c.score, found:!!parsed.ok, message:parsed.message||'', parser:'sheetjs-v50'});
    if(parsed.ok){ salary=parsed; break; }
  }
  const retireSheets = visible
    .map(s=>({s, aoa:sheetAoas.get(s.name) || sheetToAoaRobust_v50(wb.Sheets[s.name])}))
    .filter(x=>RETIRE_RE.test(x.s.name) || sheetContentHit(x.aoa, /(퇴직|적립금이월액|퇴직적립금|퇴직급여충당)/))
    .map(x=>parseRetireSheet(x.s.name, x.aoa));
  return {fileName, parser:'sheetjs-v50', visibleSheets:visible.map(s=>s.name), salaryCandidates:tried, salary, retirement:retireSheets};
}

async function parseExcelWithRawXml_v50(fileName, originalBuf){
  const rawBook = await readXlsxRawWorkbook(originalBuf);
  const visible = rawBook.sheets.filter(s=>!s.hidden);
  const aoaMap = new Map();
  for(const s of visible){ try{ aoaMap.set(s.name, await rawBook.getAoa(s.name)); }catch(e){ aoaMap.set(s.name, []); } }
  const candidates = visible.map(s=>({...s, score:sheetScore_v50(s.name, aoaMap.get(s.name)||[])})).filter(s=>s.score>0).sort((a,b)=>b.score-a.score);
  let salary=null; const tried=[];
  for(const c of candidates){
    const parsed = parseSalarySheet(c.name, aoaMap.get(c.name)||[]);
    tried.push({sheet:c.name, score:c.score, found:!!parsed.ok, message:parsed.message||'', parser:'raw-xlsx-xml-v50'});
    if(parsed.ok){ salary=parsed; break; }
  }
  const retireSheets = visible
    .filter(s=>RETIRE_RE.test(s.name) || sheetContentHit(aoaMap.get(s.name)||[], /(퇴직|적립금이월액|퇴직적립금|퇴직급여충당)/))
    .map(s=>parseRetireSheet(s.name, aoaMap.get(s.name)||[]));
  return {fileName, parser:'raw-xlsx-xml-v50', visibleSheets:visible.map(s=>s.name), salaryCandidates:tried, salary, retirement:retireSheets};
}

function sheetScore_v50(name, aoa){
  let score = sheetScore(name);
  const joined = (aoa || []).slice(0,120).map(r=>(r||[]).map(text).join(' ')).join(' ');
  const n = norm(joined);
  if(/교직원.*보수|보수.*일람|보수기준|급여기준|직명.*성명|본봉|기본급|지급액계|월지급액|소계교원|소계직원|소계일반직/.test(n)) score += 90;
  if(/직명/.test(n) && (/본봉|기본급/.test(n)) && /지급액/.test(n)) score += 100;
  if(/퇴직|적립금이월액/.test(n)) score -= 60;
  if(/세출예산명세서|세입예산명세서/.test(n)) score -= 80;
  return Math.max(0, score);
}
function sheetContentHit(aoa, re){
  const joined = (aoa || []).slice(0,100).map(r=>(r||[]).map(text).join(' ')).join(' ');
  return re.test(joined);
}

function sheetToAoaRobust_v50(ws){
  if(!ws || !ws['!ref']) return [];
  const aoa = sheetToAoaRobust(ws);
  let range;
  try{ range = XLSX.utils.decode_range(ws['!ref']); }catch(e){ return aoa; }
  const maxR = Math.min(range.e.r, 200), maxC = Math.min(range.e.c, 260);
  for(let R=range.s.r; R<=maxR; R++){
    if(!aoa[R]) aoa[R]=[];
    for(let C=range.s.c; C<=maxC; C++){
      if(aoa[R][C] !== undefined && aoa[R][C] !== '') continue;
      const addr = XLSX.utils.encode_cell({r:R,c:C}); const cell = ws[addr];
      if(!cell) continue;
      if(cell.w !== undefined && cell.w !== '') aoa[R][C]=cell.w;
      else if(cell.v !== undefined && cell.v !== '') aoa[R][C]=cell.v;
      else if(cell.f) aoa[R][C]=evaluateSimpleFormula(ws, cell.f);
    }
  }
  return aoa;
}
function evaluateSimpleFormula(ws, f){
  try{
    let expr = String(f||'').replace(/^=/,'').toUpperCase();
    expr = expr.replace(/SUM\(([A-Z]+\d+):([A-Z]+\d+)\)/g, (_,a,b)=>String(sumRange(ws,a,b)));
    expr = expr.replace(/([A-Z]+\d+)/g, m=>String(toNum(ws[m]?.v ?? ws[m]?.w ?? 0)));
    if(/^[0-9+\-*/().\s]+$/.test(expr)){
      const val = Function('"use strict";return ('+expr+')')();
      return isFinite(val) ? val : '';
    }
  }catch(e){}
  return '';
}
function sumRange(ws,a,b){
  const ca = XLSX.utils.decode_cell(a), cb = XLSX.utils.decode_cell(b);
  let s=0;
  for(let r=Math.min(ca.r,cb.r); r<=Math.max(ca.r,cb.r); r++) for(let c=Math.min(ca.c,cb.c); c<=Math.max(ca.c,cb.c); c++){
    const addr = XLSX.utils.encode_cell({r,c}); s += toNum(ws[addr]?.v ?? ws[addr]?.w ?? 0);
  }
  return s;
}

async function parseSourceFormPdf(file){
  const parsed = await parsePdf(file);
  const lines = parsed.lines || [];
  const joined = lines.map(l=>l.텍스트).join(' ');
  const retirement = parseRetireFromSourcePdfLines(lines);
  const salary = parseSalaryFromSourcePdfLines(lines);
  return {
    fileName:file.name,
    parser:'source-form-pdf-v50',
    sourcePdfNotice:SOURCE_PDF_NOTICE,
    visibleSheets:['PDF 작성서식'],
    salaryCandidates:[{sheet:'PDF 작성서식', score: salary?.ok ? 100 : 0, found:!!salary?.ok, message: salary?.ok ? 'PDF 텍스트에서 보수표 추정' : 'PDF에서 교직원보수 표를 구조적으로 읽지 못했습니다.', parser:'source-form-pdf-v50'}],
    salary,
    retirement,
    pdfLines:lines.slice(0,500),
    contentHint: joined.slice(0,1000)
  };
}

function parseRetireFromSourcePdfLines(lines){
  const out=[];
  const hitPages = new Set(lines.filter(l=>/퇴직|적립금이월액|퇴직적립금/.test(l.텍스트)).map(l=>l.페이지));
  for(const page of hitPages){
    const chunk = lines.filter(l=>l.페이지===page).map(l=>l.텍스트).join(' ');
    const nums = [...chunk.matchAll(/[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{6,}/g)].map(m=>toNum(m[0])).filter(n=>n>0);
    const amount = nums.length ? nums[nums.length-1] : 0;
    out.push({sheetName:`PDF ${page}쪽 퇴직금`, positiveAmount:amount, hasRetirementAmount:amount>0, retireColumns:[], positiveCells:amount?[{행:page, 열:'PDF', 값:amount, 헤더:'PDF 퇴직/적립금 추정', 주변:chunk.slice(0,180)}]:[], 기준:'PDF 텍스트 내 퇴직/적립금 추정'});
  }
  return out;
}

function parseSalaryFromSourcePdfLines(lines){
  // 텍스트형 작성서식 PDF에서 열 구조가 보존된 경우를 위한 1차 추정 파서입니다.
  // 정확한 엑셀 표 구조가 사라진 PDF는 결과 확인 표에서 원문을 확인해야 합니다.
  const salaryLines = lines.filter(l=>/교직원|보수|급여|수당|직명|본봉|기본급|지급액|소계/.test(l.텍스트));
  if(!salaryLines.length) return null;
  const teacherRows=[], staffRows=[], allowancesMap=new Map();
  let section='';
  for(let idx=0; idx<salaryLines.length; idx++){
    const raw = salaryLines[idx].텍스트;
    const n = norm(raw);
    if(/교원|원장|교사/.test(n) && /소계|합계|계/.test(n)) section='교원';
    if(/직원|일반직|사무|조리|차량|보조/.test(n) && /소계|합계|계/.test(n)) section='직원';
    const nums = [...raw.matchAll(/[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{6,}/g)].map(m=>toNum(m[0])).filter(x=>x>0);
    if(!nums.length) continue;
    if(/소계|합계|계/.test(n)){
      const last = nums[nums.length-1];
      const base = nums.find(x=>x>=1000000) || last;
      const row = {행:idx+1, 직명: section || (/(교원|교사|원장)/.test(n)?'교원':'직원'), 성명:'PDF소계', 본봉:base, 지급액계:last, 수당:{}};
      if(/교원|원장|교사/.test(n)) teacherRows.push(row); else staffRows.push(row);
    }
    const allowanceCandidates = ['정액급식비','급식비','명절휴가비','스승의날상여금','방학휴가비','성과상여금','시간외수당','직급보조비','관리업무수당','기타수당','자가운전보조금','연구활동비','연구수당'];
    for(const key of allowanceCandidates){
      if(norm(raw).includes(norm(key))){
        const k = cleanHeader(key); const amount = nums[nums.length-1];
        if(!allowancesMap.has(k)) allowancesMap.set(k,{항목:k, 교원금액:0, 직원금액:0, 교원인원:0, 직원인원:0});
        const obj=allowancesMap.get(k);
        if(/직원|사무|조리|차량|보조|일반직/.test(n)) obj.직원금액 += amount; else obj.교원금액 += amount;
      }
    }
  }
  const tBase = teacherRows.reduce((s,r)=>s+Number(r.본봉||0),0), sBase=staffRows.reduce((s,r)=>s+Number(r.본봉||0),0);
  if(!tBase && !sBase && !allowancesMap.size) return null;
  return {
    ok:true, sheetName:'PDF 작성서식', header:{source:'PDF 텍스트 추정', notice:SOURCE_PDF_NOTICE},
    teacherRows, staffRows,
    summary:{
      교원:{인원:teacherRows.length, 본봉:tBase, 지급액계:teacherRows.reduce((s,r)=>s+Number(r.지급액계||0),0), 소계본봉:tBase, 소계지급액계:teacherRows.reduce((s,r)=>s+Number(r.지급액계||0),0)},
      직원:{인원:staffRows.length, 본봉:sBase, 지급액계:staffRows.reduce((s,r)=>s+Number(r.지급액계||0),0), 소계본봉:sBase, 소계지급액계:staffRows.reduce((s,r)=>s+Number(r.지급액계||0),0)},
      수당:[...allowancesMap.values()].filter(x=>x.교원금액||x.직원금액)
    },
    message:'작성서식 PDF 텍스트에서 추정 추출했습니다. 원문 PDF 구조에 따라 확인이 필요합니다.'
  };
}

function getExcelRetirementSummary(retirementSheets){
  const sheets = retirementSheets || [];
  const amount = sheets.reduce((sum,r)=>sum+Number(r.positiveAmount||0),0);
  const byKind = {교원:0, 직원:0, 공통:0};
  for(const r of sheets){
    const ctx = norm([r.sheetName, ...(r.positiveCells||[]).map(c=>`${c.헤더||''} ${c.주변||''}`)].join(' '));
    const val = Number(r.positiveAmount||0);
    if(!val) continue;
    if(/교원|교사|원장/.test(ctx)) byKind.교원 += val;
    else if(/직원|일반직|사무|조리|차량|보조|영양/.test(ctx)) byKind.직원 += val;
    else byKind.공통 += val;
  }
  return {amount, has:amount>0, byKind, sheets:sheets.filter(r=>Number(r.positiveAmount||0)>0).map(r=>r.sheetName), cells:sheets.flatMap(r=>(r.positiveCells||[]).map(c=>({...c, sheetName:r.sheetName}))).slice(0,30)};
}

function getPdfRetirementSummary(pdfItems){
  const items = pdfItems || [];
  const all = items.filter(x=>Number(x.PDF금액||0)>0 && (isRetirementName(x.항목 || x.목) || isRetirementName(x.산출기초||'')));
  const totalRows = all.filter(x=>x.구분==='총액행');
  const used = totalRows.length ? totalRows : all.filter(x=>x.구분==='산출기초');
  const byName = new Map();
  for(const x of used){
    const key = norm(x.항목 || x.목 || x.산출기초);
    const prev = byName.get(key);
    if(!prev || Number(x.PDF금액||0) > Number(prev.PDF금액||0)) byName.set(key,x);
  }
  const vals=[...byName.values()];
  const byKind={교원:0, 직원:0, 공통:0};
  for(const x of vals){
    const ctx = norm(`${x.목||''} ${x.항목||''} ${x.산출기초||''}`);
    if(/교원|교사|원장/.test(ctx)) byKind.교원 += Number(x.PDF금액||0);
    else if(/직원|일반직|사무|조리|차량|보조|영양/.test(ctx)) byKind.직원 += Number(x.PDF금액||0);
    else byKind.공통 += Number(x.PDF금액||0);
  }
  return {has:vals.length>0, amount:vals.reduce((s,x)=>s+Number(x.PDF금액||0),0), items:vals, byKind};
}

function retirementVerdict(excelSummary, pdfSummary){
  const ea = Number(excelSummary.amount || 0), pa = Number(pdfSummary.amount || 0);
  if(ea > 0 && pa > 0){
    const notes=[];
    if(excelSummary.byKind?.교원>0 && !(pdfSummary.byKind?.교원>0 || pdfSummary.byKind?.공통>0)) notes.push('교원 퇴직적립금 편성 확인 필요');
    if(excelSummary.byKind?.직원>0 && !(pdfSummary.byKind?.직원>0 || pdfSummary.byKind?.공통>0)) notes.push('직원 퇴직적립금 편성 확인 필요');
    return notes.length ? `편성 일부 확인 필요(${notes.join(', ')})` : `편성 확인(금액 비교 제외 · PDF ${fmt(pa)})`;
  }
  if(ea > 0 && pa <= 0) return `미편성(작성서식 적립금액 있음 · PDF 퇴직 관련 편성 없음)`;
  if(ea <= 0 && pa > 0) return `참고: PDF에는 퇴직 관련 편성이 있으나 작성서식 적립금액은 없음(금액 비교 제외)`;
  return '해당 없음(작성서식 적립금액 없음 · PDF 퇴직 관련 편성 없음)';
}

function buildPrecheck(report){
  const rows=[]; const issues=[];
  const salaryObj = report.excel?.salary; const ex = salaryObj?.summary;
  const pdfItems = report.pdf?.items || [];
  const totals = pdfItems.filter(x=>x.구분==='총액행'); const calcs = pdfItems.filter(x=>x.구분==='산출기초');
  const totalByName = (name) => totalRowByName(totals, name)?.PDF금액 || 0;
  const addIssue=(지적내용,근거)=>issues.push({번호:issues.length+1,지적내용,근거});
  addBasisIssues(issues, calcs);
  if(ex){
    const teacherBase = ex.교원.소계본봉 || ex.교원.본봉;
    const staffBase = ex.직원.소계본봉 || ex.직원.본봉;
    const teacherPayCalcs = exactMokCalcs(calcs, '교원급여');
    const teacherPayTotal = totalByName('교원급여') || sumAmount(teacherPayCalcs);
    const teacherOff = findOffMokMatches(calcs, '교원급여', amountDiff(teacherBase, teacherPayTotal), '교원급여');
    rows.push({구분:'급여', 항목:'교원급여', 엑셀금액:teacherBase, PDF금액:teacherPayTotal, 검토결과:detailVerdict('교원급여', teacherBase, teacherPayTotal, ex.교원.인원, sumPeopleDistinct(teacherPayCalcs)||'', teacherOff)});
    const staffPayCalcs = exactMokCalcs(calcs, '직원급여');
    const staffPayTotal = totalByName('직원급여') || sumAmount(staffPayCalcs);
    const staffOff = findOffMokMatches(calcs, '직원급여', amountDiff(staffBase, staffPayTotal), '직원급여');
    rows.push({구분:'급여', 항목:'직원급여', 엑셀금액:staffBase, PDF금액:staffPayTotal, 검토결과:detailVerdict('직원급여', staffBase, staffPayTotal, ex.직원.인원, sumPeopleDistinct(staffPayCalcs)||'', staffOff)});
    const allowanceRows = ex.수당.filter(a=>!LEGAL_RE.test(a.항목));
    const teacherAnalysis = analyzeAllowances('교원', allowanceRows, pdfItems, totalRowByName(totals, '교원수당'));
    const staffAnalysis = analyzeAllowances('직원', allowanceRows, pdfItems, totalRowByName(totals, '직원수당'));
    for(const d of teacherAnalysis.details) rows.push({구분:'수당', 항목:d.항목, 엑셀금액:d.엑셀금액, PDF금액:d.PDF금액, 검토결과:d.검토결과});
    for(const d of staffAnalysis.details) rows.push({구분:'수당', 항목:d.항목, 엑셀금액:d.엑셀금액, PDF금액:d.PDF금액, 검토결과:d.검토결과});
    addOffMokDiffIssues_v50(issues, calcs, [
      {category:'교원급여', expectedMok:'교원급여', excelAmount:teacherBase, pdfAmount:teacherPayTotal},
      {category:'직원급여', expectedMok:'직원급여', excelAmount:staffBase, pdfAmount:staffPayTotal},
      {category:'교원수당', expectedMok:'교원수당', excelAmount:teacherAnalysis.totalExcel, pdfAmount:teacherAnalysis.groupMatchedAmount || teacherAnalysis.groupTotal},
      {category:'직원수당', expectedMok:'직원수당', excelAmount:staffAnalysis.totalExcel, pdfAmount:staffAnalysis.groupMatchedAmount || staffAnalysis.groupTotal}
    ]);
    addMissingAllowanceIssues(issues, teacherAnalysis, staffAnalysis);
  }
  const excelRetireSummary = getExcelRetirementSummary(report.excel?.retirement || []);
  const pdfRetireSummary = getPdfRetirementSummary(pdfItems);
  rows.push({구분:'퇴직', 항목:'퇴직적립금', 엑셀금액:excelRetireSummary.amount, PDF금액:pdfRetireSummary.amount, 검토결과:retirementVerdict(excelRetireSummary, pdfRetireSummary)});
  if(excelRetireSummary.has && !pdfRetireSummary.has) addIssue('퇴직적립금 미편성', '작성서식에는 적립금이월액 계 금액이 있으나 세출예산명세서에서 교원/직원 퇴직금및퇴직적립금 편성을 찾지 못했습니다. 금액 일치 여부는 검토하지 않습니다.');
  if(excelRetireSummary.has && pdfRetireSummary.has){
    if(excelRetireSummary.byKind?.교원>0 && !(pdfRetireSummary.byKind?.교원>0 || pdfRetireSummary.byKind?.공통>0)) addIssue('교원퇴직적립금 편성 확인 필요', '작성서식에서 교원 퇴직 적립금액이 확인되었으나 PDF에서 교원퇴직금및퇴직적립금 목이 명확히 확인되지 않았습니다.');
    if(excelRetireSummary.byKind?.직원>0 && !(pdfRetireSummary.byKind?.직원>0 || pdfRetireSummary.byKind?.공통>0)) addIssue('직원퇴직적립금 편성 확인 필요', '작성서식에서 직원 퇴직 적립금액이 확인되었으나 PDF에서 직원퇴직금및퇴직적립금 목이 명확히 확인되지 않았습니다.');
  }
  if(!excelRetireSummary.has && pdfRetireSummary.has) addIssue('작성서식 적립금액 없음 / PDF 퇴직 관련 편성 확인', '작성서식의 적립금이월액 계 금액은 확인되지 않았으나 PDF에는 퇴직 관련 편성이 있습니다. 필요 시 원자료를 확인하세요.');
  return {rows, issues, notices:[SOURCE_PDF_NOTICE]};
}

function addOffMokDiffIssues_v50(issues, calcs, checks){
  const seen = new Set(issues.map(i=>i.지적내용+'|'+i.근거));
  for(const c of checks){
    const diff = amountDiff(c.excelAmount, c.pdfAmount);
    if(closeMoney(diff,0)) continue;
    const matches = findOffMokMatches(calcs, c.expectedMok, diff, c.category);
    for(const x of matches){
      const item = cleanItemName(x.항목 || c.category);
      const detail = `${item} ${fmt(x.PDF금액)}을 ${c.expectedMok}이 아닌 ${x.목 || x.상위항목 || '다른 목'}에 편성`;
      const basis = `${c.category} 차액 ${fmt(Math.abs(diff))}과 일치. PDF ${x.페이지 || ''}쪽 산출기초: ${x.산출기초 || ''}`;
      const key = detail+'|'+basis; if(seen.has(key)) continue;
      seen.add(key); issues.push({번호:issues.length+1, 지적내용:detail, 근거:basis});
    }
  }
}

function render(report){
  if(report.excel){
    const ex=report.excel;
    $('excelSummary').innerHTML = `<span class="pill">파일 ${escapeHtml(ex.fileName)}</span><span class="pill">파서 ${escapeHtml(ex.parser||'-')}</span><span class="pill">보이는 시트 ${ex.visibleSheets?.length||0}개</span>` + (ex.sourcePdfNotice ? `<p class="muted">${escapeHtml(ex.sourcePdfNotice)}</p>` : '');
    let html = '<h3 class="section-title">시트/표 후보</h3>'+table(['sheet','score','found','message'], ex.salaryCandidates||[], {small:true});
    if(ex.salary?.ok){
      html += `<h3 class="section-title">선택된 보수 표: ${escapeHtml(ex.salary.sheetName)}</h3>`;
      html += table(['항목','값'], Object.entries(ex.salary.header||{}).map(([항목,값])=>({항목,값:JSON.stringify(값)})), {small:true});
      html += '<h3 class="section-title">보수 요약</h3>'+table(['구분','본봉','소계본봉','지급액계','소계지급액계','인원'], [
        {구분:'교원', 본봉:fmt(ex.salary.summary.교원.본봉), 소계본봉:fmt(ex.salary.summary.교원.소계본봉), 지급액계:fmt(ex.salary.summary.교원.지급액계), 소계지급액계:fmt(ex.salary.summary.교원.소계지급액계), 인원:ex.salary.summary.교원.인원},
        {구분:'직원', 본봉:fmt(ex.salary.summary.직원.본봉), 소계본봉:fmt(ex.salary.summary.직원.소계본봉), 지급액계:fmt(ex.salary.summary.직원.지급액계), 소계지급액계:fmt(ex.salary.summary.직원.소계지급액계), 인원:ex.salary.summary.직원.인원}
      ]);
      html += '<h3 class="section-title">수당 추출</h3>'+table(['항목','교원금액','직원금액'], (ex.salary.summary.수당||[]).map(a=>({항목:a.항목, 교원금액:fmt(a.교원금액), 직원금액:fmt(a.직원금액)})));
    }else html += '<p class="bad">보수 표를 구조적으로 읽지 못했습니다. 작성서식 PDF라면 스캔본 여부를 확인하고, 엑셀 원본 사용을 권장합니다.</p>';
    html += '<h3 class="section-title">퇴직 관련 표</h3>'+table(['sheetName','positiveAmount','hasRetirementAmount','positiveCells'], (ex.retirement||[]).map(r=>({sheetName:r.sheetName, positiveAmount:fmt(r.positiveAmount), hasRetirementAmount:r.hasRetirementAmount?'있음':'없음', positiveCells:(r.positiveCells||[]).map(c=>`${c.행}행 ${c.열}열 ${fmt(c.값)}`).join(' / ')})));
    if(ex.pdfLines) html += '<h3 class="section-title">작성서식 PDF 원문 라인</h3>'+table(['페이지','y','텍스트'], ex.pdfLines.slice(0,300), {small:true});
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
  let reviewHtml = '<h3 class="section-title">금액 검토</h3>' + table(['구분','항목','엑셀금액','PDF금액','검토결과'], (review.rows||[]).map(r=>({구분:r.구분, 항목:r.항목, 엑셀금액:fmt(r.엑셀금액), PDF금액:fmt(r.PDF금액), 검토결과:r.검토결과})));
  reviewHtml += '<h3 class="section-title">지적사항</h3>' + table(['번호','지적내용','근거'], review.issues || []);
  $('precheck').innerHTML = reviewHtml;
}

/* ===== v59 targeted stabilization overrides =====
   목표: 행복한아이들 정상 결과는 유지하면서
   - 해아뜰: 기타수당/상여금 복잡 산식은 개별편성으로 확정
   - 생각키움: 교원수당 내부 분할 산출내역 복구, 퇴직 총액 중복/오인식 완화,
              직원급여 차액의 같은 목 합산 오편성 탐색
*/
function isRightAllowanceMok_v59(kind, x){
  const mok = norm(x?.목 || x?.상위항목 || '');
  const allowanceMok = kind === '교원' ? '교원수당' : '직원수당';
  return mok === allowanceMok || mok.includes(allowanceMok);
}
function genericBonusCandidate_v59(kind, x){
  const label = norm(`${x?.항목 || ''} ${x?.산출기초 || ''}`);
  if(!isRightAllowanceMok_v59(kind, x)) return false;
  if(!/상여/.test(label)) return false;
  // 일반 상여금과 별도 항목을 혼동하지 않도록 제외
  if(/성과상여|스승의날|명절휴가|방학휴가/.test(label)) return false;
  return true;
}
function allowanceCandidateByKey_v59(kind, key, x){
  if(!x || x.구분 !== '산출기초') return false;
  if(LEGAL_RE.test(x.항목 || '')) return false;
  if(!isRightAllowanceMok_v59(kind, x)) return false;
  const label = norm(`${cleanItemName(x.항목 || '')} ${x.산출기초 || ''}`);
  switch(key){
    case '정액급식': return /(정액급식|급식수당|급식비교원|교원정액급식비|직원정액급식비)/.test(label) && !/(무상급식|식재료|급식운영|원아급식|교직원식재료)/.test(label);
    case '직급': return /(직급보조비|직급수당|급보조비)/.test(label);
    case '시간외': return /(연장근로수당|연장수당|시간외수당|장근로수당)/.test(label);
    case '연구': return /(연구수당|연구활동비|연구비|구수당)/.test(label) && !/연수비/.test(label);
    case '자가운전': return /(교통보조비|교통보조금|자가운전|운전수당|통보조비)/.test(label);
    case '직책급': return /직책수당|직책급/.test(label);
    case '근속': return /근속수당/.test(label);
    case '기타': return /기타수당/.test(label);
    case '상여': return genericBonusCandidate_v59(kind, x);
    case '성과상여': return /성과상여금/.test(label);
    case '스승의날상여': return /스승의날상여금/.test(label);
    case '명절휴가': return /명절휴가비/.test(label);
    case '방학휴가': return /방학휴가비/.test(label);
    default: return false;
  }
}
function findAllowancePdfMatches(kind, allowanceName, amount, pdfItems){
  const keys = allowanceKeys(allowanceName);
  const seen = new Set();
  const matches = [];
  const excelAmt = Number(amount || 0);

  const addIfOk = (x) => {
    const pdfAmt = Number(x.PDF금액 || 0);
    if(excelAmt && !(closeMoney(pdfAmt, excelAmt) || pdfAmt <= excelAmt + 1000)) return false;
    const id = `${x.페이지}|${x.행}|${x.목}|${x.항목}|${x.PDF금액}`;
    if(seen.has(id)) return false;
    seen.add(id); matches.push(x); return true;
  };

  // 1) 엄격한 목 내부 key 매칭. v56처럼 PDF 전체로 번지지 않게 막습니다.
  for(const x of (pdfItems || [])){
    if(!keys.some(key => allowanceCandidateByKey_v59(kind, key, x))) continue;
    addIfOk(x);
  }

  // 2) 기존 matcher도 보조로 사용하되, 해당 수당 목 내부로 한정합니다.
  for(const x of (pdfItems || [])){
    if(x.구분 !== '산출기초') continue;
    if(!isRightAllowanceMok_v59(kind, x)) continue;
    if(!keys.some(key => allowancePdfHit(kind, key, amount, x))) continue;
    addIfOk(x);
  }

  // 3) [방]/[급]/재원별로 쪼개진 같은 수당은 합산합니다.
  const currentSum = matches.reduce((s,x)=>s+Number(x.PDF금액||0),0);
  if(keys.length && excelAmt && currentSum && !closeMoney(currentSum, excelAmt)){
    for(const x of (pdfItems || [])){
      if(x.구분 !== '산출기초') continue;
      if(!isRightAllowanceMok_v59(kind, x)) continue;
      const id = `${x.페이지}|${x.행}|${x.목}|${x.항목}|${x.PDF금액}`;
      if(seen.has(id)) continue;
      if(!keys.some(key => allowanceCandidateByKey_v59(kind, key, x))) continue;
      const projected = matches.reduce((s,m)=>s+Number(m.PDF금액||0),0) + Number(x.PDF금액||0);
      if(projected <= excelAmt + 1000 || closeMoney(projected, excelAmt)){
        seen.add(id); matches.push(x);
        if(closeMoney(projected, excelAmt)) break;
      }
    }
  }
  return matches;
}

function retirementStrictName_v59(s){
  const n = norm(s || '');
  if(!/퇴직/.test(n) && !/직적립/.test(n)) return false;
  if(/법정부담|건강보험|국민연금|사학연금|고용보험|산재보험|일용잡급|대체/.test(n)) return false;
  return /퇴직금|퇴직적립|퇴직금및퇴직적립|직적립금/.test(n);
}
function getPdfRetirementSummary(pdfItems){
  const items = pdfItems || [];
  // 금액 비교는 하지 않지만, 표시 금액은 목 총액행을 우선해 중복 산출기초 합산을 방지합니다.
  const totalRows = items.filter(x => x.구분 === '총액행' && Number(x.PDF금액||0)>0 && retirementStrictName_v59(`${x.항목||''} ${x.목||''}`));
  const source = totalRows.length ? totalRows : items.filter(x => x.구분 === '산출기초' && Number(x.PDF금액||0)>0 && retirementStrictName_v59(`${x.목||''} ${x.항목||''} ${x.산출기초||''}`));
  const byName = new Map();
  for(const x of source){
    const key = norm(x.항목 || x.목 || x.산출기초 || '');
    if(!key) continue;
    const prev = byName.get(key);
    if(!prev || Number(x.PDF금액||0) > Number(prev.PDF금액||0)) byName.set(key, x);
  }
  const vals = [...byName.values()];
  const byKind = {교원:0, 직원:0, 공통:0};
  for(const x of vals){
    const ctx = norm(`${x.목||''} ${x.항목||''} ${x.산출기초||''}`);
    if(/교원|교사|원장/.test(ctx)) byKind.교원 += Number(x.PDF금액||0);
    else if(/직원|일반직|사무|조리|차량|보조|영양/.test(ctx)) byKind.직원 += Number(x.PDF금액||0);
    else byKind.공통 += Number(x.PDF금액||0);
  }
  return {has:vals.length>0, amount:vals.reduce((s,x)=>s+Number(x.PDF금액||0),0), items:vals, byKind};
}

function offMokAllowedPay_v59(x, categoryNorm){
  const label = norm(`${cleanItemName(x?.항목 || '')} ${x?.산출기초 || ''} ${x?.목 || x?.상위항목 || ''}`);
  if(/직원급여/.test(categoryNorm)){
    return /(차량기사급여|기사급여|통학차량|차량임차|차량운행|영양사|기본급영양사|급여|인건비)/.test(label)
           && !/(주유비|유류비|수리비|보험료|차량운영비)/.test(label);
  }
  if(/교원급여/.test(categoryNorm)) return /(교원급여|교사급여|원장급여|방과후교원급여|기본급)/.test(label);
  return /(급여|기본급|인건비|보수|영양사)/.test(label);
}
function findOffMokMatches(calcs, expectedMok, diff, category){
  const absDiff = Math.abs(Number(diff||0));
  if(absDiff <= 1000) return [];
  const target = norm(expectedMok);
  const categoryNorm = norm(category || expectedMok);
  const isPay = /급여/.test(categoryNorm);
  const candidates = (calcs || []).filter(x=>{
    if(!x || !x.PDF금액) return false;
    if(LEGAL_RE.test(x.항목 || '')) return false;
    const mok = norm(x.목 || x.상위항목 || '');
    if(!mok || mok.includes(target)) return false;
    if(isPay && !offMokAllowedPay_v59(x, categoryNorm)) return false;
    return true;
  });
  const direct = candidates
    .filter(x => closeMoney(Math.abs(Number(x.PDF금액||0)), absDiff))
    .sort((a,b)=>offMokMatchScore(b, category, expectedMok)-offMokMatchScore(a, category, expectedMok));
  if(direct.length) return direct;

  // 단일 산출내역이 아니라 같은 목 안의 여러 산출내역 합계가 차액과 일치하는 경우.
  // 예: 통학차량이용비의 통학차량임차료(수익자)+통학차량임차료 = 차량기사 급여 차액
  const byMok = new Map();
  for(const x of candidates){
    const mok = x.목 || x.상위항목 || '';
    if(!byMok.has(mok)) byMok.set(mok, []);
    byMok.get(mok).push(x);
  }
  const combos=[];
  for(const [mok, arr] of byMok){
    const useful = arr.filter(x=>Number(x.PDF금액||0)>0).sort((a,b)=>Number(b.PDF금액||0)-Number(a.PDF금액||0));
    // 전체 합계 우선
    const sum = useful.reduce((s,x)=>s+Number(x.PDF금액||0),0);
    if(closeMoney(sum, absDiff)){
      combos.push({페이지:useful[0]?.페이지, 행:useful[0]?.행, 목:mok, 상위항목:mok, 항목:useful.map(x=>cleanItemName(x.항목)).join(' + '), PDF금액:sum, 산출기초:useful.map(x=>`${cleanItemName(x.항목)} ${fmt(x.PDF금액)}`).join(' + '), 구분:'합산오편성', _combined:useful});
      continue;
    }
    // 작은 수의 부분합 탐색
    const n = Math.min(useful.length, 10);
    for(let mask=1; mask < (1<<n); mask++){
      const subset=[]; let s=0;
      for(let i=0;i<n;i++) if(mask & (1<<i)){ subset.push(useful[i]); s += Number(useful[i].PDF금액||0); }
      if(subset.length < 2) continue;
      if(closeMoney(s, absDiff)){
        combos.push({페이지:subset[0]?.페이지, 행:subset[0]?.행, 목:mok, 상위항목:mok, 항목:subset.map(x=>cleanItemName(x.항목)).join(' + '), PDF금액:s, 산출기초:subset.map(x=>`${cleanItemName(x.항목)} ${fmt(x.PDF금액)}`).join(' + '), 구분:'합산오편성', _combined:subset});
        break;
      }
    }
  }
  combos.sort((a,b)=>offMokMatchScore(b, category, expectedMok)-offMokMatchScore(a, category, expectedMok));
  return combos;
}
function detailVerdict(label, excelAmount, pdfAmount, excelPeople, pdfPeople, offMatches=[]){
  const parts=[];
  const ea=Number(excelAmount||0), pa=Number(pdfAmount||0);
  if(!closeMoney(ea, pa)){
    const diff=ea-pa;
    if(offMatches && offMatches.length){
      const m=offMatches[0];
      let offText='';
      if(m._combined && m._combined.length){
        const names = m._combined.map(x=>`${cleanItemName(x.항목)} ${fmt(x.PDF금액)}`).join(' + ');
        offText = `${m.목 || m.상위항목} 목에 ${names}으로 편성`;
      }else{
        const itemName = cleanItemName(m.항목);
        offText = (/직원급여/.test(label) && /통학차량|차량/.test(norm((m.목||'') + itemName)))
          ? `${m.목 || m.상위항목} 목의 ${itemName} ${fmt(m.PDF금액)}으로 편성`
          : `${m.목 || m.상위항목} 목에 ${itemName} ${fmt(m.PDF금액)} 편성`;
      }
      parts.push(`금액 차이(${label} ${shortWon(Math.abs(diff))} 차이 → ${offText})`);
    }else{
      parts.push(`금액 차이(엑셀 ${fmt(ea)} / PDF ${fmt(pa)} / 차이 ${fmt(Math.abs(diff))})`);
    }
  }
  return parts.length ? parts.join(' / ') : '일치';
}
