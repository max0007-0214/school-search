const CACHE='school-search-v18';
const ASSETS=['./','./index.html','./app.js?v=18','./manifest.webmanifest','./icon.svg'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS.map(u=>new Request(u,{cache:'reload'})))));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil((async()=>{for(const k of await caches.keys())if(k.startsWith('school-search-')&&k!==CACHE)await caches.delete(k);await self.clients.claim()})())});
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET'||new URL(r.url).origin!==self.location.origin)return;
e.respondWith((async()=>{try{const res=await fetch(r,{cache:'no-cache'});if(res.ok){(await caches.open(CACHE)).put(r,res.clone())}return res}catch(err){const c=await caches.match(r,{ignoreSearch:false})||await caches.match(r,{ignoreSearch:true});if(c)return c;throw err}})())});
