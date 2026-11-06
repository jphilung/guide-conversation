// La version est remplacée automatiquement par build_data.py à chaque génération.
const V="v-ee3ac6f582";
const SHELL=["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png","data.json"];
self.addEventListener("install",e=>{self.skipWaiting();e.waitUntil((async()=>{
  const c=await caches.open(V);await c.addAll(SHELL);
  const d=await (await fetch("data.json")).json();
  await Promise.allSettled(d.expressions.map(x=>c.add(x.audio))); // un audio absent n'est pas bloquant
})())});
self.addEventListener("activate",e=>e.waitUntil((async()=>{
  for(const k of await caches.keys())if(k!==V)await caches.delete(k);
  await clients.claim();
})()));
async function respond(req){
  const r=await caches.match(req,{ignoreSearch:true});
  if(!r)return fetch(req);
  const rg=req.headers.get("range");
  if(!rg)return r;
  const b=await r.blob(),m=/bytes=(\d+)-(\d*)/.exec(rg),s=+m[1],e=m[2]?+m[2]+1:b.size; // Safari exige les requêtes Range
  return new Response(b.slice(s,e),{status:206,headers:{"Content-Type":r.headers.get("Content-Type")||"audio/mpeg","Content-Range":`bytes ${s}-${e-1}/${b.size}`}});
}
self.addEventListener("fetch",e=>{if(e.request.method==="GET")e.respondWith(respond(e.request))});
