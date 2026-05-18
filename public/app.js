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
  const btn = $('runBtn');
  try{
    btn.disabled = true;
    const excelFile = $('excelFile').files[0];
    const pdfFile = $('pdfFile').files[0];
    if(!excelFile && !pdfFile) throw new Error('엑셀 또는 PDF 파일을 선택해주세요.');
    const report = { createdAt:new Date().toISOString(), excel:null, pdf:null, precheck:[], progress:[] };
    const step = async (msg) => {
      report.progress.push({time:new Date().toISOString(), message:msg});
      setStatus(msg, true);
      await new Promise(r=>setTimeout(r, 20));
    };
    await step('1/4 파일 확인 완료');
    if(excelFile){
      await step('2/4 엑셀 시트와 보수/퇴직 자료를 읽는 중입니다...');
      report.excel = await parseExcel(excelFile);
      await step('2/4 엑셀 읽기 완료');
    }
    if(pdfFile){
      await step('3/4 PDF 목/산출기초를 읽는 중입니다...');
      report.pdf = await parsePdf(pdfFile);
      await step('3/4 PDF 읽기 완료');
    }
    await step('4/4 1차 검토표를 만드는 중입니다...');
    report.precheck = buildPrecheck(report);
    state.report = report;
    render(report);
    $('downloadBtn').disabled = false;
    setStatus('값 추출이 완료되었습니다. 아래 표에서 엑셀/PDF가 제대로 읽혔는지 확인해주세요.', true);
  }catch(e){ console.error(e); setStatus('오류가 발생했습니다: '+e.message, false); }
  finally{ btn.disabled = false; }
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
function parseSalarySheet(sheetName, aoa){
  const preview = aoa.slice(0,60).map((r,i)=>({행:i+1, 내용:r.map(text).filter(Boolean).join(' | ').slice(0,300)})).filter(x=>x.내용);
  if(!aoa.length) return {ok:false, sheetName, message:'빈 시트', preview};
  const maxRows = Math.min(100, aoa.length), maxCols = Math.max(...aoa.slice(0,maxRows).map(r=>r.length),0);
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
  // v39: 일부 기관 파일은 브라우저 파서가 헤더 텍스트를 부분적으로 놓칩니다.
  // 이때 소계(교원)/소계(직원) 행의 숫자 구조로 본봉·지급액계 열을 추론합니다.
  const inferred = detectSalaryHeaderBySubtotalPattern(aoa, maxRows, maxCols);
  if(inferred.ok) return inferred;
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


function detectSalaryHeaderBySubtotalPattern(aoa, maxRows, maxCols){
  const teacherEnd = findRow(aoa, 0, /소계\s*\(?교원\)?|소계교원/);
  const staffEnd = findRow(aoa, teacherEnd >= 0 ? teacherEnd + 1 : 0, /소계\s*\(?(직원|일반직)\)?|소계직원|소계일반직/);
  if(teacherEnd < 0 && staffEnd < 0) return {ok:false};
  const subtotalRow = aoa[teacherEnd >= 0 ? teacherEnd : staffEnd] || [];
  // 직명/성명은 대부분 앞쪽 텍스트 열입니다. 그래도 텍스트 밀도가 가장 높은 앞쪽 열을 찾습니다.
  let jobCol = -1, nameCol = -1;
  for(let c=0; c<Math.min(8, maxCols); c++){
    const top = columnText(aoa, c, Math.min(maxRows, 30));
    const nt = norm(top);
    if(jobCol < 0 && (/직명/.test(nt) || /원장|교사|조리|영양|보조|기사|직원/.test(top))) jobCol = c;
    if(nameCol < 0 && (/성명/.test(nt) || /[가-힣]{2,4}/.test(top) andFalse())) nameCol = c;
  }
  if(jobCol < 0) jobCol = 1;
  if(nameCol < 0) nameCol = jobCol + 1;

  const nums = [];
  for(let c=0; c<maxCols; c++){
    const n = toNum(subtotalRow[c]);
    if(n > 0) nums.push({c,n});
  }
  if(!nums.length) return {ok:false};
  const firstMoney = nums.find(x => x.c > nameCol) || nums[0];
  const baseCol = firstMoney.c;
  let totalCol = -1, bestDiff = Infinity;
  for(const cand of nums){
    const c = cand.c;
    if(c <= baseCol) continue;
    let sum = 0;
    for(let k=baseCol; k<c; k++) sum += toNum(subtotalRow[k]);
    const diff = Math.abs(sum - cand.n);
    // 지급액계는 보통 본봉~수당 합과 거의 같습니다.
    if(sum > 0 && diff < bestDiff){ bestDiff = diff; totalCol = c; }
  }
  if(totalCol < 0){
    // 합산관계가 안 잡히면 세금/공제 열 앞의 큰 금액 열을 우선 사용
    const candidates = nums.filter(x=>x.c>baseCol).sort((a,b)=>b.n-a.n);
    totalCol = candidates[0]?.c ?? -1;
  }
  if(totalCol <= baseCol) return {ok:false};
  let headerRow = findHeaderRow(aoa, [jobCol, baseCol, totalCol]);
  if(!Number.isFinite(headerRow) || headerRow < 0) headerRow = Math.max(0, Math.min(teacherEnd, staffEnd >= 0 ? staffEnd : teacherEnd) - 1);
  return {ok:true, headerRow, jobCol, nameCol, baseCol, totalCol, score:88, inferred:true};
}
function andFalse(){ return false; }

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
  // v40: 퇴직 적립 여부는 퇴직 관련 시트 안의 `기적립현황` 구역에서
  // `소계` 열에 실제 양수 금액이 있는지를 기준으로 판단합니다.
  // 다른 구역의 합계/예상액/이월액/지급액은 퇴직 적립금액으로 보지 않습니다.
  const preview = aoa.slice(0,40).map((r,i)=>({행:i+1, 내용:(r||[]).map(text).filter(Boolean).join(' | ').slice(0,240)})).filter(x=>x.내용);
  const maxRows = Math.min(aoa.length, 250);
  const maxCols = Math.max(0, ...aoa.slice(0, Math.min(maxRows, aoa.length)).map(r=>Math.min(r.length || 0, 120)));

  const markerRows = [];
  for(let r=0; r<maxRows; r++){
    const rowText = (aoa[r]||[]).map(text).join(' ');
    if(/기\s*적\s*립\s*현\s*황/.test(rowText) || /기적립현황/.test(norm(rowText))){
      markerRows.push(r);
    }
  }

  const positiveCells = [];
  const retireColumns = [];

  for(const marker of markerRows){
    const headerStart = marker;
    const headerEnd = Math.min(marker + 8, maxRows - 1);
    const dataStart = headerEnd + 1;

    // `기적립현황` 표 안의 소계 열 찾기. 병합/다중행 헤더를 고려해 마커행~8행 아래까지 열별 텍스트를 봅니다.
    const subtotalCols = [];
    for(let c=0; c<maxCols; c++){
      const h = [];
      for(let r=headerStart; r<=headerEnd; r++){
        const v = text(aoa[r]?.[c]);
        if(v) h.push(v);
      }
      const joined = h.join(' ');
      const n = norm(joined);
      if(/소계/.test(joined) || /소계/.test(n)){
        // 지급현황/예상액/부담금 쪽 소계가 같이 있을 수 있어 기적립현황 마커 근처의 열을 우선 사용합니다.
        subtotalCols.push({col:c, header:joined.slice(0,80)});
      }
    }

    for(const sc of subtotalCols){
      if(!retireColumns.some(x=>x.col===sc.col)) retireColumns.push(sc);
      let blankRun = 0;
      for(let r=dataStart; r<maxRows; r++){
        const row = aoa[r] || [];
        const rowText = row.map(text).join(' ');
        if(/퇴직\s*예상|지급\s*현황|지급액|예상액|이월|예금\s*이자|비고/.test(rowText) && r > dataStart + 2) break;
        if(!rowText){
          blankRun++;
          if(blankRun >= 4) break;
          continue;
        }
        blankRun = 0;
        const n = toNum(row[sc.col]);
        if(n > 0){
          positiveCells.push({행:r+1, 열:sc.col+1, 값:n, 헤더:sc.header, 구역:'기적립현황', 주변:rowText.slice(0,180)});
        }
      }
    }
  }

  // 소계 열에 합계행이 있으면 합계행만 사용해 상세행+합계행 중복 집계를 방지합니다.
  const totalLike = positiveCells.filter(c=>/(^|\s)(계|합계|총계|소계)(\s|$)/.test(c.주변));
  const usedCells = totalLike.length ? totalLike : positiveCells;
  const positive = usedCells.reduce((s,c)=>s+Number(c.값||0),0);

  return {
    sheetName,
    positiveAmount:positive,
    hasRetirementAmount:positive>0,
    criterion:'기적립현황 소계 열의 양수 금액',
    markerRows:markerRows.map(r=>r+1),
    retireColumns:retireColumns.map(c=>({열:c.col+1, 헤더:c.header})),
    positiveCells:usedCells.slice(0,30),
    ignoredPositiveCells:totalLike.length ? positiveCells.filter(c=>!totalLike.includes(c)).slice(0,10) : [],
    preview
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
      const name = findCalcName(lines, i) || currentTotal?.항목 || '산출항목미상';
      out.push({페이지:l.page, 행:i+1, 목:currentTotal?.항목 || '', 상위항목:currentTotal?.항목 || '', 항목:name, PDF금액:calc.amount, PDF금액천원:Math.round(calc.amount/1000), 인원:calc.people, 산출기초:t, 구분:'산출기초'});
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
  if(pdfPeople !== '' && pdfPeople !== null && pdfPeople !== undefined && ep !== pp){
    const d=ep-pp;
    const dir=d>0 ? '부족' : '초과';
    parts.push(`인원 차이(엑셀 ${ep}명 / PDF ${pp}명, ${label.replace(/급여|수당/g,'')} ${Math.abs(d)}명 ${dir})`);
  }
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
  // 소속/직종 접두어는 같은 항목 매칭에서 제외합니다.
  k = k.replace(/^(교원|직원|교사|사무직원|조리직원|보조교사)/,'');

  // 기관마다 다른 명칭을 같은 수당으로 봅니다.
  if(/^(식대|급식비|정액급식비|정액급식|급식수당)$/.test(k) || /정액급식/.test(k)) return '정액급식';
  if(/성과상여/.test(k)) return '성과상여';
  if(/명절휴가/.test(k)) return '명절휴가';
  if(/스승의날/.test(k)) return '스승의날상여';
  if(/방학휴가/.test(k)) return '방학휴가';
  if(/자가운전|자가차량|차량유지/.test(k)) return '자가운전';
  if(/직책급|직책/.test(k)) return '직책급';

  return k.replace(/수당|보조금|지원비|지원금|상여금|휴가비|급식비|식대|비/g,'');
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
    let html = '<h3 class="section-title">PDF 추출 항목</h3>'+table(['구분','페이지','목','항목','PDF금액','PDF금액천원','인원','산출기초'], pdf.items.map(x=>({...x,PDF금액:fmt(x.PDF금액)})));
    html += '<h3 class="section-title">PDF 원문 라인</h3>'+table(['페이지','y','텍스트'], pdf.lines.slice(0,300), {small:true});
    $('pdfTables').innerHTML = html;
  }
  const review = report.precheck || {rows:[], issues:[]};
  let reviewHtml = '<h3 class="section-title">금액 및 인원 검토</h3>' + table(['구분','항목','엑셀금액','엑셀인원','PDF금액','PDF인원','검토결과'], review.rows.map(r=>({...r,엑셀금액:fmt(r.엑셀금액),PDF금액:fmt(r.PDF금액)})));
  reviewHtml += '<h3 class="section-title">지적사항</h3>' + table(['번호','지적내용','근거'], review.issues || []);
  $('precheck').innerHTML = reviewHtml;
}
