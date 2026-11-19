'use strict';
const JS_VER='14';
let db;
const $=id=>document.getElementById(id);
const escapeHtml=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function showTab(t){if(t==='ask')refreshAskDocs();for(const x of ['docs','events','help','ask']){$(x).classList.toggle('hide',x!==t);$('tab'+x[0].toUpperCase()+x.slice(1)).classList.toggle('active',x===t)}if(t==='help')storageInfo();if(t==='ask')showApiState()}
function init(){return new Promise((resolve,reject)=>{const r=indexedDB.open('school-search-local-v1',1);r.onupgradeneeded=()=>{const d=r.result;if(!d.objectStoreNames.contains('docs'))d.createObjectStore('docs',{keyPath:'id'});if(!d.objectStoreNames.contains('events'))d.createObjectStore('events',{keyPath:'id'})};r.onsuccess=()=>{db=r.result;resolve()};r.onerror=()=>reject(r.error)})}
function store(name,mode='readonly'){return db.transaction(name,mode).objectStore(name)}
function getAll(name){return new Promise((resolve,reject)=>{const r=store(name).getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)})}
function put(name,obj){return new Promise((resolve,reject)=>{const r=store(name,'readwrite').put(obj);r.onsuccess=()=>resolve();r.onerror=()=>reject(r.error)})}
function del(name,id){return new Promise((resolve,reject)=>{const r=store(name,'readwrite').delete(id);r.onsuccess=()=>resolve();r.onerror=()=>reject(r.error)})}
const decoder=new TextDecoder();
function u16(v,p){return v.getUint16(p,true)}function u32(v,p){return v.getUint32(p,true)}
async function unzip(buffer){const v=new DataView(buffer),bytes=new Uint8Array(buffer),entries={};let end=-1;for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--){if(u32(v,p)===0x06054b50){end=p;break}}if(end<0)throw Error('ZIPの終端が見つかりません');const n=u16(v,end+10);let p=u32(v,end+16);for(let i=0;i<n;i++){if(u32(v,p)!==0x02014b50)throw Error('ZIPの目次が壊れています');const method=u16(v,p+10),size=u32(v,p+20),nameLen=u16(v,p+28),extra=u16(v,p+30),comment=u16(v,p+32),off=u32(v,p+42);const name=decoder.decode(bytes.slice(p+46,p+46+nameLen));p+=46+nameLen+extra+comment;if(!/^(word\/document\.xml|xl\/sharedStrings\.xml|xl\/worksheets\/sheet\d+\.xml|xl\/workbook\.xml)$/.test(name))continue;if(u32(v,off)!==0x04034b50)continue;const start=off+30+u16(v,off+26)+u16(v,off+28);let part=bytes.slice(start,start+size);if(method===8){if(!('DecompressionStream' in window))throw Error('このSafariではWord/Excelの展開に対応していません');part=new Uint8Array(await new Response(new Blob([part]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())}else if(method!==0)continue;entries[name]=decoder.decode(part)}return entries}
async function extract(file){const ext=file.name.split('.').pop().toLowerCase();if(['txt','csv','md','html','htm','json','xml'].includes(ext))return {text:await file.text(),note:''};if(['doc','xls'].includes(ext))return {text:'',note:'旧形式の本文抽出には未対応'};if(ext==='pdf')return extractPdf(await file.arrayBuffer());if(ext==='docx'||ext==='xlsx'){const e=await unzip(await file.arrayBuffer());if(ext==='docx'){if(!e['word/document.xml'])throw Error('Word本文がありません');const d=new DOMParser().parseFromString(e['word/document.xml'],'application/xml');return {text:[...d.getElementsByTagName('*')].filter(n=>n.localName==='p').map(p=>p.textContent).join('\n'),note:''}}const strings=[];if(e['xl/sharedStrings.xml']){const d=new DOMParser().parseFromString(e['xl/sharedStrings.xml'],'application/xml');strings.push(...[...d.getElementsByTagName('*')].filter(n=>n.localName==='si').map(n=>n.textContent))}const lines=[];for(const key of Object.keys(e).filter(k=>k.startsWith('xl/worksheets/sheet'))){const d=new DOMParser().parseFromString(e[key],'application/xml');lines.push(key);for(const c of [...d.getElementsByTagName('*')].filter(n=>n.localName==='c')){const val=[...c.getElementsByTagName('*')].find(n=>n.localName==='v'||n.localName==='is');if(val){let text=val.textContent||'';if(c.getAttribute('t')==='s')text=strings[Number(text)]??text;lines.push((c.getAttribute('r')||'')+': '+text)}}}return {text:lines.join('\n'),note:''}}return {text:'',note:'本文抽出非対応'}}
let pdfLibPromise;
function loadPdfLibrary(){if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);if(!pdfLibPromise)pdfLibPromise=new Promise((resolve,reject)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';s.onload=()=>window.pdfjsLib?resolve(window.pdfjsLib):reject(Error('PDF解析プログラムが見つかりません'));s.onerror=()=>reject(Error('PDF解析プログラムを取得できません。初回はインターネット接続が必要です'));document.head.appendChild(s)}).catch(e=>{pdfLibPromise=null;throw e});return pdfLibPromise}
async function readPdfOnce(lib,buffer,worker){const opts={data:new Uint8Array(buffer.slice(0)),useSystemFonts:true,disableFontFace:false,cMapUrl:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',cMapPacked:true,standardFontDataUrl:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/standard_fonts/'};if(!worker)opts.disableWorker=true;const task=lib.getDocument(opts);let pdf;try{pdf=await task.promise;let pages=[],chars=0,itemsTotal=0,nonemptyPages=0,imageOps=0,errors=[];for(let n=1;n<=pdf.numPages;n++){const page=await pdf.getPage(n);let content;try{content=await page.getTextContent({disableCombineTextItems:false,includeMarkedContent:false})}catch(e){errors.push(`${n}ページ:${e.message}`);continue}itemsTotal+=content.items.length;let parts=[],lastY=null;for(const item of content.items){if(typeof item.str!=='string')continue;const y=item.transform?.[5];if(lastY!==null&&typeof y==='number'&&Math.abs(y-lastY)>3)parts.push('\n');parts.push(item.str);if(item.hasEOL)parts.push('\n');if(typeof y==='number')lastY=y}const line=parts.join('');const count=line.replace(/\s/g,'').length;chars+=count;if(count)nonemptyPages++;pages.push(`【${n}ページ】\n${line}`);if(n<=3&&count===0){try{const ops=await page.getOperatorList();imageOps+=ops.fnArray.filter(op=>[lib.OPS.paintImageXObject,lib.OPS.paintInlineImageXObject,lib.OPS.paintJpegXObject,lib.OPS.paintImageMaskXObject].includes(op)).length}catch(e){errors.push(`画像判定:${e.message}`)}}page.cleanup()}return {text:pages.join('\n\n'),pages:pdf.numPages,chars,itemsTotal,nonemptyPages,imageOps,errors}}finally{if(pdf)await pdf.destroy();else await task.destroy().catch(()=>{})}}
// OCRは端末内のWeb Workerで実行。PDF本文・画像を外部APIへ送信しない。
let ocrLibraryPromise;
function loadOcrLibrary(){
 if(window.Tesseract)return Promise.resolve(window.Tesseract);
 if(!ocrLibraryPromise)ocrLibraryPromise=new Promise((resolve,reject)=>{
  const s=document.createElement('script');
  s.src='https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
  s.onload=()=>window.Tesseract?resolve(window.Tesseract):reject(Error('OCRプログラムを初期化できません'));
  s.onerror=()=>reject(Error('OCRプログラムを読み込めません。通信環境を確認してください'));
  document.head.appendChild(s);
 }).catch(e=>{ocrLibraryPromise=null;throw e});
 return ocrLibraryPromise;
}
function importProgress(message){$('importStatus').textContent=message}
let ocrCancel=false;
async function createOcrWorker(){
 const tess=await loadOcrLibrary();
 const logger=m=>{if(m.status==='recognizing text'&&typeof m.progress==='number'&&window.__ocrPage)importProgress(`日本語OCR：${window.__ocrPage}ページ・${Math.round(m.progress*100)}%`)};
 const base={workerPath:'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',corePath:'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1',logger};
 try{return await tess.createWorker('jpn',1,{...base,langPath:'https://cdn.jsdelivr.net/npm/@tesseract.js-data/jpn@1.0.0/4.0.0_best_int'})}
 catch(e){return await tess.createWorker('jpn',1,{...base,langPath:'https://tessdata.projectnaptha.com/4.0.0'})}
}
// 日本語文字どうしの間に入った空白を除去（検索ヒット用）
function joinCjk(t){const re=/([\u3040-\u30ff\u3400-\u9fff\uff00-\uffef])[ \t\u3000]+(?=[\u3040-\u30ff\u3400-\u9fff\uff00-\uffef])/g;let p;do{p=t;t=t.replace(re,'$1')}while(t!==p);return t}
// グレースケール化とコントラスト伸長
function enhance(ctx,w,h){const im=ctx.getImageData(0,0,w,h),d=im.data,hist=new Uint32Array(256);for(let i=0;i<d.length;i+=4){const g=(d[i]*299+d[i+1]*587+d[i+2]*114)/1000|0;d[i]=g;hist[g]++}const tot=w*h;let lo=0,hi=255,c=0;for(let i=0;i<256;i++){c+=hist[i];if(c>=tot*0.01){lo=i;break}}c=0;for(let i=255;i>=0;i--){c+=hist[i];if(c>=tot*0.01){hi=i;break}}const rg=Math.max(40,hi-lo);for(let i=0;i<d.length;i+=4){let v=((d[i]-lo)*255/rg);v=v<0?0:v>255?255:v;d[i]=d[i+1]=d[i+2]=v}ctx.putImageData(im,0,0)}
function cancelButton(show){
 let b=document.getElementById('cancelOcr');
 if(!b){b=document.createElement('button');b.id='cancelOcr';b.className='small danger';b.textContent='読み込みを中止';b.onclick=()=>{ocrCancel=true;b.textContent='中止しています…'};$('importStatus').after(b)}
 b.style.display=show?'':'none';if(show)b.textContent='読み込みを中止';
}
// ページ単位のハイブリッド抽出：文字が30字未満のページだけOCR（端末内Web Worker。本文・画像は外部送信しない）
async function extractPdf(buffer){
 const MIN_CHARS=30,RESET_EVERY=20;
 const lib=await loadPdfLibrary();
 lib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
 const mk=extra=>({data:new Uint8Array(buffer.slice(0)),useSystemFonts:true,cMapUrl:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',cMapPacked:true,standardFontDataUrl:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/standard_fonts/',...extra});
 let task=lib.getDocument(mk({})),pdf;
 try{pdf=await task.promise}catch(e){await task.destroy().catch(()=>{});task=lib.getDocument(mk({disableWorker:true}));try{pdf=await task.promise}catch(e2){throw Error('PDF解析失敗：'+e2.message)}}
 let worker=null,sinceReset=0,textPages=0,ocrPages=0,chars=0,failed=[];const out=[];
 ocrCancel=false;cancelButton(true);
 try{
  for(let n=1;n<=pdf.numPages;n++){
   if(ocrCancel)throw Error('中止しました。この資料は保存していません');
   let page,canvas;
   try{
    page=await pdf.getPage(n);
    let txt='';
    try{const c=await page.getTextContent();const ps=[];let lastY=null;for(const it of c.items){if(typeof it.str!=='string')continue;const y=it.transform?.[5];if(lastY!==null&&typeof y==='number'&&Math.abs(y-lastY)>3)ps.push('\n');ps.push(it.str);if(it.hasEOL)ps.push('\n');if(typeof y==='number')lastY=y}txt=ps.join('')}catch(e){}
    const cnt=txt.replace(/\s/g,'').length;
    let pageText=txt;
    if(cnt>=MIN_CHARS){textPages++;chars+=cnt}
    else{
     window.__ocrPage=`${n}/${pdf.numPages}`;
     importProgress(`日本語OCR：${n}/${pdf.numPages}ページ。画面を閉じないでください。`);
     try{
      if(worker&&sinceReset>=RESET_EVERY){await worker.terminate().catch(()=>{});worker=null;sinceReset=0}
      if(!worker){importProgress('日本語OCRを準備中…初回は認識データの取得に時間がかかります。');worker=await createOcrWorker();sinceReset=0}
      const o=page.getViewport({scale:1});
      const sc=Math.max(1.0,Math.min(3.5,Math.sqrt(4500000/(o.width*o.height))));
      const vp=page.getViewport({scale:sc});
      canvas=document.createElement('canvas');canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);
      const ctx=canvas.getContext('2d',{alpha:false});if(!ctx)throw Error('描画領域を確保できません');
      await page.render({canvasContext:ctx,viewport:vp,background:'#ffffff'}).promise;
      try{enhance(ctx,canvas.width,canvas.height)}catch(e){}
      const r=await worker.recognize(canvas);const ot=joinCjk(r.data?.text||'');
      const oc=ot.replace(/\s/g,'').length;sinceReset++;
      if(oc){ocrPages++;chars+=oc;pageText=(txt?txt+'\n':'')+ot}else if(!cnt)failed.push(n);
      if(cnt)chars+=cnt;
     }catch(e){if(ocrCancel)throw e;failed.push(n);if(cnt)chars+=cnt;if(worker){await worker.terminate().catch(()=>{});worker=null}}
    }
    out.push(`【${n}ページ】\n${pageText}`);
   }finally{if(canvas){canvas.width=0;canvas.height=0}if(page)page.cleanup()}
   await new Promise(r=>setTimeout(r,0));
  }
  const note=`PDF ${pdf.numPages}ページ：文字抽出${textPages}ページ・OCR${ocrPages}ページ`+(failed.length?`・認識できず${failed.length}ページ(${failed.slice(0,10).join(',')}${failed.length>10?'…':''})`:'')+`。OCRは誤認識があります`;
  return {text:out.join('\n\n'),note};
 }finally{
  window.__ocrPage=null;cancelButton(false);
  if(worker)await worker.terminate().catch(()=>{});
  await pdf.destroy().catch(()=>{});
 }
}
function garbledPdf(d){return d.name?.toLowerCase().endsWith('.pdf')&&(d.note||'').includes('簡易対応')}
async function markOldPdfs(){for(const d of await getAll('docs')){if(garbledPdf(d)){d.text='';d.note='旧版の簡易解析PDFです。原本を再登録してください。';await put('docs',d)}}}
$('files').addEventListener('change',async e=>{const files=[...e.target.files];let ok=0,fail=[];$('importStatus').textContent='読み込み中…';for(const f of files){try{const r=await extract(f);if(f.name.toLowerCase().endsWith('.pdf')&&!r.text.replace(/【\d+ページ】/g,'').trim())throw Error('PDF本文が空です');await put('docs',{id:f.name+'|'+f.size+'|'+f.lastModified,name:f.name,size:f.size,modified:f.lastModified,added:Date.now(),text:r.text,note:r.note});ok++}catch(err){fail.push(f.name+'：'+err.message)}}$('importStatus').textContent=`${ok}件保存しました。`+(fail.length?' 失敗: '+fail.join(' / '):'');e.target.value='';renderDocs()});
function fuzzyFind(hay,w){const L=w.length,maxMis=L>=7?2:L>=4?1:0;if(!maxMis)return -1;for(let i=0;i+L<=hay.length;i++){let m=0;for(let j=0;j<L;j++){if(hay.charCodeAt(i+j)!==w.charCodeAt(j)&&++m>maxMis)break}if(m<=maxMis)return i}return -1}
async function renderDocs(){if(!db)return;
 const norm=x=>String(x||'').normalize('NFKC').toLowerCase().replace(/\s+/g,'');
 const ws=$('query').value.trim().split(/\s+/).map(norm).filter(Boolean),docs=await getAll('docs');docs.sort((a,b)=>b.added-a.added);
 let exact=[],fuzzy=[];
 if(!ws.length)exact=docs;
 else for(const d of docs){const hay=norm(d.name+d.text);let ok=true,all=true;for(const w of ws){if(hay.includes(w))continue;all=false;if(fuzzyFind(hay,w)<0){ok=false;break}}if(ok)(all?exact:fuzzy).push(d)}
 const snip=d=>{if(!ws.length)return(d.text||'').slice(0,180);const tx=norm(d.text);let ix=-1,len=0;for(const w of ws){let i=tx.indexOf(w);if(i<0)i=fuzzyFind(tx,w);if(i>=0){ix=i;len=w.length;break}}if(ix<0)return(d.text||'').slice(0,180);return tx.slice(Math.max(0,ix-40),ix)+'【'+tx.slice(ix,ix+len)+'】'+tx.slice(ix+len,ix+len+120)};
 const card=d=>`<div class="item"><div class="name">${escapeHtml(d.name)}</div><div class="meta">${new Date(d.added).toLocaleString('ja-JP')} · ${Math.ceil(d.size/1024)}KB ${d.note?'<span class="pill">抽出情報あり</span>':''}</div><div class="snippet">${escapeHtml(snip(d)||d.note||'本文なし')}</div><div class="row"><button class="small" data-action="view" data-id="${escapeHtml(d.id)}">本文を見る</button><button class="small danger" data-action="delete" data-id="${escapeHtml(d.id)}">削除</button></div></div>`;
 const total=exact.length+fuzzy.length;
 const head=!ws.length?`${docs.length}件表示 / 保存済み${docs.length}件`:total?`<b>${total}件ヒット</b>（完全一致${exact.length}件・あいまい一致${fuzzy.length}件）/ 保存済み${docs.length}件`:`<b>ヒットなし</b>（保存済み${docs.length}件）`;
 $('docsList').innerHTML=`<p class="hint">${head}</p>`+exact.map(card).join('')+(fuzzy.length?`<p class="hint">▼ あいまい一致（OCRの誤字を含む可能性。【】が一致した箇所）</p>`+fuzzy.map(card).join(''):'');
}
$('docsList').addEventListener('click',async e=>{const b=e.target.closest('button[data-action]');if(!b)return;const id=b.dataset.id;if(b.dataset.action==='delete'){if(confirm('この資料を削除しますか？')){await del('docs',id);renderDocs()}}else{const r=store('docs').get(id);r.onsuccess=()=>{const d=r.result;if(d)alert((d.text||d.note||'本文なし').slice(0,12000))}}});
async function addEvent(){const date=$('eventDate').value,title=$('eventTitle').value.trim();if(!date||!title){$('eventStatus').textContent='日付と予定名を入力してください';return}await put('events',{id:crypto.randomUUID(),date,time:$('eventTime').value,title,note:$('eventNote').value});$('eventStatus').textContent='保存しました';$('eventTitle').value='';$('eventNote').value='';renderEvents()}
async function renderEvents(){if(!db)return;const q=$('eventQuery').value.trim().toLowerCase(),all=await getAll('events');all.sort((a,b)=>(a.date+a.time).localeCompare(b.date+b.time));$('eventsList').innerHTML=all.filter(x=>(x.date+' '+x.title+' '+x.note).toLowerCase().includes(q)).map(x=>`<div class="item"><div class="name">${escapeHtml(x.date)} ${escapeHtml(x.time)}　${escapeHtml(x.title)}</div><div class="snippet">${escapeHtml(x.note)}</div><div class="row"><button class="small" data-action="calendar" data-id="${escapeHtml(x.id)}">カレンダーに追加</button><button class="small danger" data-action="delete" data-id="${escapeHtml(x.id)}">削除</button></div></div>`).join('')||'<p class="hint">予定はありません。</p>'}
$('eventsList').addEventListener('click',async e=>{const b=e.target.closest('button[data-action]');if(!b)return;const r=store('events').get(b.dataset.id);r.onsuccess=async()=>{const x=r.result;if(!x)return;if(b.dataset.action==='delete'){if(confirm('予定を削除しますか？')){await del('events',x.id);renderEvents()}}else calendarFile(x)}});
function download(name,content,type){const url=URL.createObjectURL(new Blob([content],{type})),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),3000)}
function calendarFile(x){const stamp=x.date.replaceAll('-','')+'T'+(x.time||'09:00').replace(':','')+'00',start=new Date(x.date+'T'+(x.time||'09:00')+':00'),end=new Date(start.getTime()+3600000),fmt=d=>{const z=n=>String(n).padStart(2,'0');return d.getFullYear()+z(d.getMonth()+1)+z(d.getDate())+'T'+z(d.getHours())+z(d.getMinutes())+'00'},esc=s=>String(s||'').replace(/\\/g,'\\\\').replace(/\n/g,'\\n').replace(/,/g,'\\,').replace(/;/g,'\\;');download('予定-'+x.date+'.ics',`BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//SchoolSearch//JP\r\nBEGIN:VEVENT\r\nUID:${x.id}@school-search.local\r\nDTSTAMP:${new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d+/,'')}\r\nDTSTART:${stamp}\r\nDTEND:${fmt(end)}\r\nSUMMARY:${esc(x.title)}\r\nDESCRIPTION:${esc(x.note)}\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`,'text/calendar')}
async function exportBackup(){download('校内資料検索バックアップ-'+new Date().toISOString().slice(0,10)+'.json',JSON.stringify({version:1,docs:await getAll('docs'),events:await getAll('events')}),'application/json')}
$('restore').addEventListener('change',async e=>{try{if(!e.target.files.length)return;const data=JSON.parse(await e.target.files[0].text());if(data.version!==1||!Array.isArray(data.docs)||!Array.isArray(data.events))throw Error('対応していないバックアップです');if(!confirm(`資料${data.docs.length}件・予定${data.events.length}件を復元しますか？`))return;for(const d of data.docs)await put('docs',d);for(const x of data.events)await put('events',x);alert('復元しました');renderDocs();renderEvents()}catch(err){alert('復元失敗: '+err.message)}e.target.value=''});
async function storageInfo(){const docs=await getAll('docs'),events=await getAll('events');const bytes=new Blob([JSON.stringify({docs,events})]).size;const estimate=await navigator.storage?.estimate?.().catch(()=>null);$('storageStatus').textContent=`資料 ${docs.length}件、予定 ${events.length}件。保存データ概算 ${(bytes/1048576).toFixed(2)}MB`+(estimate?`／このサイトの使用量概算 ${(estimate.usage/1048576).toFixed(1)}MB`:'');}
init().then(async()=>{await markOldPdfs();renderDocs();renderEvents();$('eventDate').value=new Date().toLocaleDateString('sv-SE');if('serviceWorker' in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./sw.js').catch(()=>{})}).catch(e=>alert('端末内保存を利用できません: '+e.message));

// ===== v14: 資料から予定を抽出／資料への質問 =====
// ==PURE-START==
const nk=s=>String(s||'').normalize('NFKC');
function fiscalYearOf(doc){
 const t=nk(doc.text||'');let m=t.match(/令和\s*(元|\d{1,2})\s*年度/);
 if(m)return 2018+(m[1]==='元'?1:+m[1]);
 m=t.match(/(20\d\d)\s*年度/);if(m)return +m[1];
 m=nk(doc.name).match(/^[Rr](\d{1,2})(\d{2})(\d{2})(?!\d)/);
 if(m){const y=2018+ +m[1],mo=+m[2];return mo>=4?y:y-1}
 return null;
}
function extractDates(doc){
 const text=nk(doc.text||''),now=new Date();
 const fy=fiscalYearOf(doc)??(now.getMonth()>=3?now.getFullYear():now.getFullYear()-1);
 const WD='日月火水木金土',out=[],seen=new Set();let page=0;
 const isD=c=>c>='0'&&c<='9';
 for(let line of text.split('\n')){
  const pm=line.match(/^【(\d+)ページ】/);if(pm){page=+pm[1];line=line.slice(pm[0].length)}
  if(!line.trim())continue;
  const used=new Array(line.length).fill(false),found=[];
  const pats=[
   {re:/令和\s*(元|\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/g,f:m=>[2018+(m[1]==='元'?1:+m[1]),+m[2],+m[3]],conf:'high'},
   {re:/(20\d{2})\s*[年\/\-.]\s*(\d{1,2})\s*[月\/\-.]\s*(\d{1,2})\s*日?/g,f:m=>[+m[1],+m[2],+m[3]],conf:'high'},
   {re:/(\d{1,2})\s*月\s*(\d{1,2})\s*日/g,f:m=>{const mo=+m[1];return[mo>=4?fy:fy+1,mo,+m[2]]},conf:'high'},
   {re:/(\d{1,2})\s*\/\s*(\d{1,2})/g,f:m=>{const mo=+m[1];return[mo>=4?fy:fy+1,mo,+m[2]]},conf:'low'}
  ];
  for(const p of pats){
   for(const m of line.matchAll(p.re)){
    const s=m.index,e=s+m[0].length;
    if(s>0&&(isD(line[s-1])||(p.conf==='low'&&line[s-1]==='/')))continue;
    if(e<line.length&&(isD(line[e])||(p.conf==='low'&&line[e]==='/')))continue;
    if(used.slice(s,e).some(Boolean))continue;
    const [y,mo,d]=p.f(m);
    if(mo<1||mo>12||d<1||d>31)continue;
    const dt=new Date(y,mo-1,d);if(dt.getMonth()!==mo-1)continue;
    for(let i=s;i<e;i++)used[i]=true;
    const wm=line.slice(e,e+10).match(/^\s*[（(]\s*([日月火水木金土])\s*[)）]?/);
    const wdv=wm?(WD[dt.getDay()]===wm[1]?'ok':'ng'):'';found.push({s,e,raw:m[0],wmLen:wm?wm[0].length:0,y,mo,d,conf:(p.conf==='low'&&wdv==='ok')?'high':p.conf,wd:wdv});
   }
  }
  if(!found.length)continue;
  let tm='';const t1=line.match(/(\d{1,2}):(\d{2})/)||line.match(/(\d{1,2})時\s*(\d{2})\s*分/)||line.match(/(\d{1,2})時(?!間)/);
  if(t1){const h=+t1[1],mi=t1[2]?+t1[2]:0;if(h<=23&&mi<=59)tm=String(h).padStart(2,'0')+':'+String(mi).padStart(2,'0')}
  let title=line;for(const f of [...found].sort((a,b)=>b.s-a.s))title=title.slice(0,f.s)+' '+title.slice(f.e+f.wmLen);
  title=title.replace(/[\s　]+/g,' ').replace(/^[\s・:：,、\-~〜～]+|[\s・:：,、\-~〜～]+$/g,'').slice(0,30);
  for(const f of found){
   const iso=`${f.y}-${String(f.mo).padStart(2,'0')}-${String(f.d).padStart(2,'0')}`,key=iso+'|'+line.slice(0,50);
   if(seen.has(key))continue;seen.add(key);
   out.push({date:iso,time:tm,title,line:line.trim().slice(0,100),page,conf:f.conf,wd:f.wd});
  }
 }
 out.sort((a,b)=>a.date<b.date?-1:a.date>b.date?1:0);
 return out.slice(0,300);
}
function chunkDoc(doc,size=450){
 const out=[];let page=0,buf='',bufPage=0;
 const flush=()=>{if(buf.trim())out.push({doc:doc.name,id:doc.id,page:bufPage,text:buf.trim()});buf=''};
 for(let ln of nk(doc.text||'').split('\n')){
  const m=ln.match(/^【(\d+)ページ】/);if(m){flush();page=+m[1];ln=ln.slice(m[0].length)}
  while(ln.length>size*1.5){flush();bufPage=page;buf=ln.slice(0,size);flush();ln=ln.slice(size)}
  if(!buf)bufPage=page;
  if(buf.length+ln.length>size){flush();bufPage=page}
  buf+=ln+'\n';
 }
 flush();return out;
}
function qBigrams(q){
 let s=nk(q).toLowerCase().replace(/(を?教えて(ください|下さい)?|でしょうか|ですか|ください|について|という|はいつ|いつ|どこ|なに|何)/g,'').replace(/[\s、。，．,.!?？！「」『』()\[\]【】・:;\-~〜～]/g,'');
 const set=new Set(),hira=/^[\u3040-\u309f]+$/;
 for(let i=0;i+1<s.length;i++){const b=s.slice(i,i+2);if(!hira.test(b))set.add(b)}
 return set;
}
function pickChunks(q,docs,k=8){
 const bg=qBigrams(q);if(!bg.size)return [];
 const scored=[];
 for(const d of docs)for(const c of chunkDoc(d)){
  const t=c.text.replace(/\s+/g,'');let sc=0;for(const b of bg)if(t.includes(b))sc++;
  if(sc>0)scored.push({...c,score:sc});
 }
 scored.sort((a,b)=>b.score-a.score);
 const top=scored[0]?.score||0;
 return scored.filter(c=>c.score>=Math.max(2,top*0.4)).slice(0,k);
}
function buildPrompt(q,chunks){
 const system='あなたは学校の資料を読むアシスタントです。以下の「抜粋」だけを根拠に、質問に日本語で簡潔に答えてください。\n・抜粋に書かれていないことは「資料からは分かりません」と答えてください。\n・抜粋はOCRで読み取った文字で、誤字・脱字を含みます。日付や数字が不自然なときは、その旨を指摘してください。\n・回答の最後に、根拠にした資料名とページ番号を示してください。';
 const ex=chunks.map((c,i)=>`【抜粋${i+1}】資料：${c.doc}　ページ：${c.page||'不明'}\n${c.text}`).join('\n\n');
 return {system,user:`${ex}\n\n【質問】\n${q}`};
}
// ==PURE-END==

function refreshAskDocs(){getAll('docs').then(ds=>{ds=ds.filter(d=>(d.text||'').trim());ds.sort((a,b)=>b.added-a.added);$('dateDoc').innerHTML=ds.length?ds.map(d=>`<option value="${escapeHtml(d.id)}">${escapeHtml(d.name)}</option>`).join(''):'<option value="">本文のある資料がありません</option>'})}
let dateCands=[],dateCandDoc=null;
async function runDateExtract(){
 const id=$('dateDoc').value;if(!id){$('dateList').innerHTML='<p class="hint">資料を選んでください。</p>';return}
 const d=(await getAll('docs')).find(x=>x.id===id);if(!d)return;
 dateCandDoc=d;dateCands=extractDates(d);
 if(!dateCands.length){$('dateList').innerHTML='<p class="hint">日付らしい表記が見つかりませんでした。</p>';return}
 const wdn=['日','月','火','水','木','金','土'];
 $('dateList').innerHTML=`<p class="hint"><b>${dateCands.length}件の候補</b>。内容を確認・修正して、登録する予定にチェックを付けてください。月/日だけの表記と、曜日が合わない候補は、既定ではチェックしていません。</p>`+dateCands.map((c,i)=>{
  const dt=new Date(c.date+'T00:00:00'),flag=c.wd==='ng'?'<span class="pill">⚠曜日が合いません（OCR誤りの可能性）</span>':c.wd==='ok'?'<span class="pill" style="background:#27432f;color:#b9f0c6">曜日一致</span>':'',low=c.conf==='low'?'<span class="pill">月/日形式・要確認</span>':'';
  const on=c.conf==='high'&&c.wd!=='ng'?'checked':'';
  return `<div class="item" data-i="${i}"><label style="display:flex;gap:10px;align-items:center"><input type="checkbox" data-f="on" style="width:auto" ${on}><span class="name">${c.date}（${wdn[dt.getDay()]}）</span></label><div class="meta">${c.page?c.page+'ページ ':''}${flag} ${low}</div><div class="snippet" style="max-height:80px">${escapeHtml(c.line)}</div><label>予定名</label><input data-f="title" value="${escapeHtml(c.title||'')}" placeholder="予定名を入力"><div class="row"><div><label>日付</label><input type="date" data-f="date" value="${c.date}"></div><div><label>時刻（任意）</label><input type="time" data-f="time" value="${c.time}"></div></div></div>`}).join('')+'<button class="primary" style="margin-top:12px;width:100%" onclick="registerDates()">チェックした予定を登録</button><div id="dateStatus" class="status"></div>';
}
async function registerDates(){
 const ex=await getAll('events'),has=new Set(ex.map(e=>e.date+'|'+e.title));let ok=0,dup=0,bad=0;
 for(const el of $('dateList').querySelectorAll('.item')){
  if(!el.querySelector('[data-f=on]').checked)continue;
  const date=el.querySelector('[data-f=date]').value,title=el.querySelector('[data-f=title]').value.trim(),time=el.querySelector('[data-f=time]').value,c=dateCands[+el.dataset.i];
  if(!date||!title){bad++;continue}
  if(has.has(date+'|'+title)){dup++;continue}
  await put('events',{id:crypto.randomUUID(),date,time,title,note:`資料：${dateCandDoc?.name||''}${c&&c.page?' p.'+c.page:''}（自動抽出）`});has.add(date+'|'+title);ok++;
 }
 $('dateStatus').textContent=`${ok}件を登録しました。`+(dup?` 同じ日付・予定名の${dup}件は重複のため登録していません。`:'')+(bad?` 日付または予定名が空の${bad}件は登録していません。`:'');
 renderEvents();
}
let askPrompt=null;
async function runAskPrepare(){
 const q=$('askQ').value.trim();askPrompt=null;$('askAnswer').textContent='';$('askActions').classList.add('hide');$('askPromptBox').classList.add('hide');
 if(!q){$('askExcerpts').innerHTML='<p class="hint">質問を入力してください。</p>';return}
 const docs=(await getAll('docs')).filter(d=>(d.text||'').trim());
 const picks=pickChunks(q,docs);
 if(!picks.length){$('askExcerpts').innerHTML='<p class="hint">関連する箇所が見つかりませんでした。質問の言い方を変えるか、資料検索でキーワードを確認してください。</p>';return}
 window.__picks=picks;
 $('askExcerpts').innerHTML=`<p class="hint"><b>${picks.length}件の抜粋</b>を見つけました。送らない抜粋は、チェックを外してください。</p>`+picks.map((c,i)=>`<div class="item"><label style="display:flex;gap:10px;align-items:center"><input type="checkbox" data-i="${i}" style="width:auto" checked onchange="updateAskPrompt()"><span class="name">${escapeHtml(c.doc)}${c.page?'　'+c.page+'ページ':''}</span></label><div class="snippet" style="max-height:110px">${escapeHtml(c.text)}</div></div>`).join('');
 $('askActions').classList.remove('hide');updateAskPrompt();
}
function updateAskPrompt(){
 const q=$('askQ').value.trim(),sel=[...$('askExcerpts').querySelectorAll('input[data-i]')].filter(x=>x.checked).map(x=>window.__picks[+x.dataset.i]);
 askPrompt=sel.length?buildPrompt(q,sel):null;
 const box=$('askPromptBox');
 if(askPrompt){box.value=askPrompt.system+'\n\n'+askPrompt.user;box.classList.remove('hide');$('askSize').textContent=`送る内容：抜粋${sel.length}件・約${(askPrompt.system+askPrompt.user).length}文字`}
 else{box.classList.add('hide');$('askSize').textContent='抜粋が選ばれていません'}
}
async function copyAskPrompt(){
 if(!askPrompt)return;const t=$('askPromptBox').value;
 try{await navigator.clipboard.writeText(t);$('askAnswer').textContent='コピーしました。Claudeアプリのチャットに貼り付けて送信してください。'}
 catch(e){const b=$('askPromptBox');b.focus();b.select();$('askAnswer').textContent='自動コピーできませんでした。上の枠の文字を選択してコピーしてください。'}
}
const monthKey=()=>'ss_api_cnt_'+new Date().toISOString().slice(0,7);
function showApiState(){
 const k=localStorage.getItem('ss_api_key')||'';
 $('apiState').textContent=(k?`キー保存済み（末尾 ${k.slice(-4)}）`:'キー未設定')+` ／ 今月の送信回数：${localStorage.getItem(monthKey())||0}回`;
 $('apiModel').value=localStorage.getItem('ss_api_model')||'claude-haiku-5-5';
}
function saveApiSettings(){
 const k=$('apiKey').value.trim(),m=$('apiModel').value.trim();
 if(k)localStorage.setItem('ss_api_key',k);
 if(m)localStorage.setItem('ss_api_model',m);
 $('apiKey').value='';showApiState();
}
function clearApiKey(){localStorage.removeItem('ss_api_key');showApiState()}
async function askApi(){
 const key=localStorage.getItem('ss_api_key')||'';
 if(!key){$('askAnswer').textContent='先に下の「AI接続の設定」でAPIキーを保存してください。';return}
 if(!askPrompt){$('askAnswer').textContent='抜粋が選ばれていません。';return}
 const n=[...$('askExcerpts').querySelectorAll('input[data-i]')].filter(x=>x.checked).length,names=[...new Set([...$('askExcerpts').querySelectorAll('input[data-i]')].filter(x=>x.checked).map(x=>window.__picks[+x.dataset.i].doc))].join('、');
 if(!confirm(`次の内容を、AnthropicのAPIへ送信します。\n\n・抜粋：${n}件（約${(askPrompt.system+askPrompt.user).length}文字）\n・資料：${names}\n\n個人情報を含む資料は、この機能に使わない運用になっていることを確認してください。送信しますか？`))return;
 $('askAnswer').textContent='回答を待っています…';
 const ac=new AbortController(),to=setTimeout(()=>ac.abort(),60000);
 try{
  const res=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',signal:ac.signal,headers:{'content-type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},body:JSON.stringify({model:localStorage.getItem('ss_api_model')||'claude-haiku-5-5',max_tokens:1000,system:askPrompt.system,messages:[{role:'user',content:askPrompt.user}]})});
  const data=await res.json().catch(()=>({}));
  if(!res.ok){const msg=data?.error?.message||('エラー '+res.status);$('askAnswer').textContent=(res.status===401?'APIキーが正しくありません。':res.status===429?'利用制限に達しました。しばらくしてからやり直してください。':'送信できませんでした。')+'\n'+msg;return}
  const txt=(data.content||[]).filter(x=>x.type==='text').map(x=>x.text).join('\n');
  $('askAnswer').textContent=txt||'（回答が空でした）';
  localStorage.setItem(monthKey(),String((+localStorage.getItem(monthKey())||0)+1));showApiState();
 }catch(e){$('askAnswer').textContent=e.name==='AbortError'?'60秒以内に回答がなく、中止しました。':'通信できませんでした：'+e.message}
 finally{clearTimeout(to)}
}

async function versionDiag(){
 const el=document.getElementById('verDiag');if(!el)return;
 let sw='SW未使用',cache='なし';
 try{
  if('serviceWorker' in navigator){const r=await navigator.serviceWorker.getRegistration();sw=r&&r.active?'SW有効':'SW未登録'}
  if(window.caches){const k=(await caches.keys()).filter(x=>x.startsWith('school-search-'));cache=k.join(',')||'なし'}
 }catch(e){}
 const html=window.__HTML_VER||'?';
 const ok=html===JS_VER&&cache==='school-search-v14';
 el.textContent=`画面v${html} / JS v${JS_VER} / ${sw} / キャッシュ ${cache} → `+(ok?'✅ 最新(v14)':'⚠ 不一致：一度アプリを完全に閉じて開き直してください');
}
versionDiag();
