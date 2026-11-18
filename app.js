'use strict';
let db;
const $=id=>document.getElementById(id);
const escapeHtml=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function showTab(t){for(const x of ['docs','events','help']){$(x).classList.toggle('hide',x!==t);$('tab'+x[0].toUpperCase()+x.slice(1)).classList.toggle('active',x===t)}if(t==='help')storageInfo()}
function init(){return new Promise((resolve,reject)=>{const r=indexedDB.open('school-search-local-v1',1);r.onupgradeneeded=()=>{let d=r.result;d.createObjectStore('docs',{keyPath:'id'});d.createObjectStore('events',{keyPath:'id'})};r.onsuccess=()=>{db=r.result;resolve()};r.onerror=()=>reject(r.error)})}
function store(name,mode='readonly'){return db.transaction(name,mode).objectStore(name)}
function getAll(name){return new Promise((resolve,reject)=>{let r=store(name).getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)})}
function put(name,obj){return new Promise((resolve,reject)=>{let r=store(name,'readwrite').put(obj);r.onsuccess=resolve;r.onerror=()=>reject(r.error)})}
function del(name,id){return new Promise((resolve,reject)=>{let r=store(name,'readwrite').delete(id);r.onsuccess=resolve;r.onerror=()=>reject(r.error)})}
const decoder=new TextDecoder();
function u16(v,p){return v.getUint16(p,true)}function u32(v,p){return v.getUint32(p,true)}
async function unzip(buffer){let v=new DataView(buffer),bytes=new Uint8Array(buffer),entries={};let end=-1;for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--){if(u32(v,p)===0x06054b50){end=p;break}}if(end<0)throw Error('ZIPの終端が見つかりません');let n=u16(v,end+10),p=u32(v,end+16);for(let i=0;i<n;i++){if(u32(v,p)!==0x02014b50)break;let method=u16(v,p+10),size=u32(v,p+20),nameLen=u16(v,p+28),extra=u16(v,p+30),comment=u16(v,p+32),off=u32(v,p+42);let name=decoder.decode(bytes.slice(p+46,p+46+nameLen));p+=46+nameLen+extra+comment;if(!/^(word\/document\.xml|xl\/sharedStrings\.xml|xl\/worksheets\/sheet\d+\.xml|xl\/workbook\.xml)$/.test(name))continue;if(u32(v,off)!==0x04034b50)continue;let start=off+30+u16(v,off+26)+u16(v,off+28),part=bytes.slice(start,start+size);if(method===8){if(!('DecompressionStream' in window))throw Error('このSafariは圧縮ファイル展開に対応していません');let stream=new Blob([part]).stream().pipeThrough(new DecompressionStream('deflate-raw'));part=new Uint8Array(await new Response(stream).arrayBuffer())}else if(method!==0)continue;entries[name]=decoder.decode(part)}return entries}
function xmlText(xml){const d=new DOMParser().parseFromString(xml,'application/xml');return d.documentElement.textContent||''}
async function extract(file){let ext=file.name.split('.').pop().toLowerCase();if(['txt','csv','md','html','htm','json','xml'].includes(ext))return {text:await file.text(),note:''};if(['doc','xls'].includes(ext))return {text:'',note:'旧形式の本文抽出には未対応'};if(ext==='docx'||ext==='xlsx'){let e=await unzip(await file.arrayBuffer());if(ext==='docx'){let xml=e['word/document.xml'];if(!xml)throw Error('Word本文がありません');let d=new DOMParser().parseFromString(xml,'application/xml');return {text:[...d.getElementsByTagName('*')].filter(n=>n.localName==='p').map(p=>p.textContent).join('\n'),note:''}}let strings=[];if(e['xl/sharedStrings.xml']){let d=new DOMParser().parseFromString(e['xl/sharedStrings.xml'],'application/xml');strings=[...d.getElementsByTagName('*')].filter(n=>n.localName==='si').map(n=>n.textContent)}let lines=[];for(let key of Object.keys(e).filter(k=>k.startsWith('xl/worksheets/sheet'))){let d=new DOMParser().parseFromString(e[key],'application/xml');let cells=[...d.getElementsByTagName('*')].filter(n=>n.localName==='c');lines.push(key);for(let c of cells){let val=[...c.getElementsByTagName('*')].find(n=>n.localName==='v'||n.localName==='is');if(val){let text=val.textContent||'';if(c.getAttribute('t')==='s')text=strings[Number(text)]??text;lines.push((c.getAttribute('r')||'')+': '+text)}}}return {text:lines.join('\n'),note:''}}
if(ext==='pdf')return extractPdf(await file.arrayBuffer());return {text:'',note:'本文抽出非対応'}}
let pdfLibPromise;
function loadPdfLibrary(){
 if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);
 if(!pdfLibPromise)pdfLibPromise=new Promise((resolve,reject)=>{
  const script=document.createElement('script');
  script.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  script.onload=()=>window.pdfjsLib?resolve(window.pdfjsLib):reject(Error('PDF解析ライブラリを読み込めません'));
  script.onerror=()=>reject(Error('PDF解析ライブラリの取得に失敗しました。インターネット接続を確認してください'));
  document.head.appendChild(script);
 }).catch(err=>{pdfLibPromise=null;throw err});
 return pdfLibPromise;
}
async function extractPdf(buffer){
 const lib=await loadPdfLibrary();
 lib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
 const task=lib.getDocument({data:new Uint8Array(buffer),useSystemFonts:true});
 const pdf=await task.promise;
 let pages=[],chars=0;
 try{
  for(let n=1;n<=pdf.numPages;n++){
   const page=await pdf.getPage(n);
   const content=await page.getTextContent();
   let parts=[],lastY=null;
   for(const item of content.items){
    if(typeof item.str!=='string')continue;
    const y=item.transform?.[5];
    if(lastY!==null&&typeof y==='number'&&Math.abs(y-lastY)>3)parts.push('\n');
    parts.push(item.str);
    if(item.hasEOL)parts.push('\n');
    if(typeof y==='number')lastY=y;
   }
   const line=parts.join('');chars+=line.replace(/\s/g,'').length;
   pages.push('【'+n+'ページ】\n'+line);
  }
 }finally{await pdf.destroy()}
 return {text:pages.join('\n\n'),note:chars===0?'文字データを検出できませんでした。画像PDFにはOCRが必要です。':''};
}
function garbledPdf(d){return d.name?.toLowerCase().endsWith('.pdf') && (d.note||'').includes('簡易対応')}
async function markOldPdfs(){let docs=await getAll('docs');for(const d of docs){if(garbledPdf(d)){d.text='';d.note='旧版でPDFを誤解析したため本文を非表示にしました。日本語検索は修正待ちです。';await put('docs',d)}}}
$('files').addEventListener('change',async e=>{let files=[...e.target.files],ok=0,fail=[];$('importStatus').textContent='読み込み中… PDFはページ数により時間がかかります。';for(let f of files){try{let result=await extract(f);await put('docs',{id:f.name+'|'+f.size+'|'+f.lastModified,name:f.name,size:f.size,modified:f.lastModified,added:Date.now(),text:result.text,note:result.note});ok++}catch(err){fail.push(f.name+'：'+err.message)}}$('importStatus').textContent=`${ok}件保存しました。`+(fail.length?' 失敗: '+fail.join(' / '):'');e.target.value='';renderDocs()});
async function renderDocs(){if(!db)return;let q=$('query').value.trim().toLowerCase().split(/\s+/).filter(Boolean),docs=await getAll('docs');docs.sort((a,b)=>b.added-a.added);let found=docs.filter(d=>q.every(w=>(d.name+' '+d.text).toLowerCase().includes(w)));$('docsList').innerHTML=`<p class="hint">${found.length}件表示 / 保存済み${docs.length}件</p>`+found.map(d=>{let ix=q.length?d.text.toLowerCase().indexOf(q[0]):-1;let snippet=ix>=0?d.text.slice(Math.max(0,ix-65),ix+150):d.text.slice(0,180);return `<div class="item"><div class="name">${escapeHtml(d.name)}</div><div class="meta">${new Date(d.added).toLocaleString('ja-JP')} · ${Math.ceil(d.size/1024)}KB ${d.note?'<span class="pill">抽出制限あり</span>':''}</div><div class="snippet">${escapeHtml(snippet||d.note||'本文を抽出できませんでした')}</div><div class="row"><button class="small" data-action="view" data-id="${escapeHtml(d.id)}">本文を見る</button><button class="small danger" data-action="delete" data-id="${escapeHtml(d.id)}">削除</button></div></div>`}).join('')}
$('docsList').addEventListener('click',async e=>{let b=e.target.closest('button[data-action]');if(!b)return;let id=b.dataset.id;if(b.dataset.action==='delete'){if(confirm('この資料を削除しますか？')){await del('docs',id);renderDocs()}}else{let r=store('docs').get(id);r.onsuccess=()=>{let d=r.result;if(d)alert((d.text||d.note||'本文なし').slice(0,12000))}}});
async function addEvent(){let date=$('eventDate').value,title=$('eventTitle').value.trim();if(!date||!title){$('eventStatus').textContent='日付と予定名を入力してください';return}await put('events',{id:crypto.randomUUID(),date,time:$('eventTime').value,title,note:$('eventNote').value});$('eventStatus').textContent='保存しました';$('eventTitle').value='';$('eventNote').value='';renderEvents()}
async function renderEvents(){if(!db)return;let q=$('eventQuery').value.trim().toLowerCase(),all=await getAll('events');all.sort((a,b)=>(a.date+a.time).localeCompare(b.date+b.time));$('eventsList').innerHTML=all.filter(x=>(x.date+' '+x.title+' '+x.note).toLowerCase().includes(q)).map(x=>`<div class="item"><div class="name">${escapeHtml(x.date)} ${escapeHtml(x.time)}　${escapeHtml(x.title)}</div><div class="snippet">${escapeHtml(x.note)}</div><div class="row"><button class="small" data-action="calendar" data-id="${x.id}">カレンダーに追加</button><button class="small danger" data-action="delete" data-id="${x.id}">削除</button></div></div>`).join('')||'<p class="hint">予定はありません。</p>'}
$('eventsList').addEventListener('click',async e=>{let b=e.target.closest('button[data-action]');if(!b)return;let r=store('events').get(b.dataset.id);r.onsuccess=async()=>{let x=r.result;if(!x)return;if(b.dataset.action==='delete'){if(confirm('予定を削除しますか？')){await del('events',x.id);renderEvents()}}else calendarFile(x)}});
function download(name,content,type){let url=URL.createObjectURL(new Blob([content],{type})),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),3000)}
function calendarFile(x){let stamp=(x.date.replaceAll('-','')+'T'+(x.time||'09:00').replace(':','')+'00'),start=new Date(x.date+'T'+(x.time||'09:00')+':00'),end=new Date(start.getTime()+3600000),fmt=d=>{let z=n=>String(n).padStart(2,'0');return d.getFullYear()+z(d.getMonth()+1)+z(d.getDate())+'T'+z(d.getHours())+z(d.getMinutes())+'00'};let esc=s=>String(s||'').replace(/\\/g,'\\\\').replace(/\n/g,'\\n').replace(/,/g,'\\,').replace(/;/g,'\\;');download('予定-'+x.date+'.ics',`BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//SchoolSearch//JP\r\nBEGIN:VEVENT\r\nUID:${x.id}@school-search.local\r\nDTSTAMP:${new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d+/,'')}\r\nDTSTART:${stamp}\r\nDTEND:${fmt(end)}\r\nSUMMARY:${esc(x.title)}\r\nDESCRIPTION:${esc(x.note)}\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`,'text/calendar')}
async function exportBackup(){download('校内資料検索バックアップ-'+new Date().toISOString().slice(0,10)+'.json',JSON.stringify({version:1,docs:await getAll('docs'),events:await getAll('events')}),'application/json')}
$('restore').addEventListener('change',async e=>{try{let data=JSON.parse(await e.target.files[0].text());if(data.version!==1||!Array.isArray(data.docs)||!Array.isArray(data.events))throw Error('対応していないバックアップです');if(!confirm(`資料${data.docs.length}件・予定${data.events.length}件を復元しますか？`))return;for(let d of data.docs)await put('docs',d);for(let x of data.events)await put('events',x);alert('復元しました');renderDocs();renderEvents()}catch(err){alert('復元失敗: '+err.message)}e.target.value=''});
async function storageInfo(){$('storageStatus').textContent=`資料 ${(await getAll('docs')).length}件、予定 ${(await getAll('events')).length}件を保存中。`}
init().then(async()=>{await markOldPdfs();renderDocs();renderEvents();$('eventDate').value=new Date().toLocaleDateString('sv-SE');if('serviceWorker' in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./sw.js').catch(()=>{})}).catch(e=>alert('端末内保存を利用できません: '+e.message));
