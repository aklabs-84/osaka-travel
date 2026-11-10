/* 오사카 여행 플래너 — 오프라인 도우미 (한 번 열어 두면 인터넷이 약해도 열려요) */
const V='osk-shell-v7',RT='osk-rt-v1';
const SHELL=['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png','./audio/p0.mp3','./audio/p1.mp3','./audio/p2.mp3','./audio/p3.mp3','./audio/p4.mp3','./audio/p5.mp3','./audio/p6.mp3','./audio/p7.mp3','./audio/p8.mp3','./audio/p9.mp3','./audio/p10.mp3','./audio/p11.mp3','./audio/p12.mp3','./audio/p13.mp3','./audio/p14.mp3','./audio/p15.mp3','./audio/p16.mp3','./audio/p17.mp3','./audio/p18.mp3','./audio/p19.mp3','./audio/p20.mp3','./audio/p21.mp3','./audio/p22.mp3','./audio/p23.mp3','./audio/p24.mp3','./audio/p25.mp3','./audio/p26.mp3','./audio/p27.mp3','./audio/p28.mp3','./audio/p29.mp3','./audio/p30.mp3','./audio/p31.mp3','./audio/p32.mp3','./audio/p33.mp3','./audio/p34.mp3','./audio/p35.mp3','./audio/p36.mp3','./audio/p37.mp3','./audio/p38.mp3','./audio/p39.mp3'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V&&k!==RT).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
async function trim(c,max){const ks=await c.keys();if(ks.length>max)for(const k of ks.slice(0,ks.length-max))await c.delete(k)}
self.addEventListener('fetch',e=>{
  const r=e.request;
  if(r.method!=='GET')return;
  const u=new URL(r.url);
  if(u.hostname.includes('nominatim'))return;               // 검색은 항상 새로
  if(u.origin===location.origin){                           // 내 앱 파일: 새 것 먼저, 안 되면 저장본
    e.respondWith(fetch(r).then(res=>{const cp=res.clone();caches.open(V).then(c=>c.put(r,cp));return res}).catch(()=>caches.match(r).then(m=>m||caches.match('./index.html'))));
    return;
  }
  e.respondWith(caches.open(RT).then(async c=>{            // 글꼴·지도 조각: 저장본 먼저, 뒤에서 새로고침
    const m=await c.match(r);
    const net=fetch(r).then(res=>{if(res&&(res.ok||res.type==='opaque')){c.put(r,res.clone());trim(c,1500)}return res}).catch(()=>m);
    return m||net;
  }));
});
