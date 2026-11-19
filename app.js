'use strict';
const JS_VER='13';
let db;
const $=id=>document.getElementById(id);
const escapeHtml=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function showTab(t){for(const x of ['docs','events','help']){$(x).classList.toggle('hide',x!==t);$('tab'+x[0].toUpperCase()+x.slice(1)).classList.toggle('active',x===t)}if(t==='help')storageInfo()}
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

async function versionDiag(){
 const el=document.getElementById('verDiag');if(!el)return;
 let sw='SW未使用',cache='なし';
 try{
  if('serviceWorker' in navigator){const r=await navigator.serviceWorker.getRegistration();sw=r&&r.active?'SW有効':'SW未登録'}
  if(window.caches){const k=(await caches.keys()).filter(x=>x.startsWith('school-search-'));cache=k.join(',')||'なし'}
 }catch(e){}
 const html=window.__HTML_VER||'?';
 const ok=html===JS_VER&&cache==='school-search-v13';
 el.textContent=`画面v${html} / JS v${JS_VER} / ${sw} / キャッシュ ${cache} → `+(ok?'✅ 最新(v13)':'⚠ 不一致：一度アプリを完全に閉じて開き直してください');
}
versionDiag();
