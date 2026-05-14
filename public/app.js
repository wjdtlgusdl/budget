
const debug = document.getElementById('debug');
document.getElementById('runBtn').addEventListener('click', async ()=>{
 const file = document.getElementById('excel').files[0];
 if(!file){ alert('엑셀 업로드'); return; }
 const buf = await file.arrayBuffer();
 const wb = XLSX.read(buf,{type:'array'});
 debug.textContent = JSON.stringify(wb.SheetNames,null,2);
});
