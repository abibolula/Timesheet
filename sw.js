/* Cache only the application shell. Google Sheets responses and submissions
   always use the network. Increment VERSION when shell assets change.
   Updates activate after old app windows close; never reload an unsaved form. */
const VERSION = 'v2';
const PREFIX = 'timekeeper-' + encodeURIComponent(self.registration.scope) + '-';
const CACHE = PREFIX + VERSION;
const ASSETS = ['index.html','manifest.webmanifest','icons/icon-192.png','icons/icon-512.png','icons/icon-maskable-512.png','icons/apple-touch-icon.png','icons/logo.svg','icons/favicon-16.png','icons/favicon-32.png','favicon.ico'].map(path=>new URL(path,self.registration.scope).href);
self.addEventListener('install',event=>event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(ASSETS))));
self.addEventListener('activate',event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(key=>key.startsWith(PREFIX)&&key!==CACHE).map(key=>caches.delete(key))))));
self.addEventListener('fetch',event=>{
 const req=event.request,url=new URL(req.url),scope=new URL(self.registration.scope);
 if(req.method!=='GET'||url.origin!==scope.origin)return;
 const isEntry=req.mode==='navigate'&&(url.pathname===scope.pathname||url.pathname===new URL('index.html',scope).pathname);
 const asset=ASSETS.find(path=>new URL(path).pathname===url.pathname);
 if(!isEntry&&!asset)return;
 const key=isEntry?ASSETS[0]:asset;
 event.respondWith((async()=>{
  const cache=await caches.open(CACHE);
  try{
   const response=await fetch(req);
   if(response.ok){await cache.put(key,response.clone());return response;}
   return (await cache.match(key))||response;
  }catch(error){const cached=await cache.match(key);if(cached)return cached;throw error;}
 })());
});
