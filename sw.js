var V='cantina-v1';
var SHELL=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png','exceljs.min.js'];
self.addEventListener('install',function(e){e.waitUntil(caches.open(V).then(function(c){return c.addAll(SHELL);}).then(function(){return self.skipWaiting();}));});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==V;}).map(function(k){return caches.delete(k);}));}).then(function(){return self.clients.claim();}));});
self.addEventListener('fetch',function(e){
  var r=e.request; if(r.method!=='GET')return;
  var u=new URL(r.url);
  if(u.origin!==location.origin&&u.hostname!=='fonts.googleapis.com'&&u.hostname!=='fonts.gstatic.com')return;
  e.respondWith(caches.open(V).then(function(c){return c.match(r).then(function(hit){
    var net=fetch(r).then(function(resp){if(resp&&resp.ok)c.put(r,resp.clone());return resp;}).catch(function(){return hit||(r.mode==='navigate'?c.match('index.html'):undefined);});
    return hit||net;});}));
});
