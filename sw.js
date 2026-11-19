/* Offline support: the app files are cached; GitHub data is never cached here
   (the app keeps the last data it loaded in its own storage). */
const V = 67, CACHE = 'mymoney-v' + V;   // bump V together with ?v= in index.html
const FILES = ['./', 'index.html', 'receipts.html', 'styles.css', 'app.js', 'styles.css?v=' + V, 'app.js?v=' + V, 'manifest.webmanifest', 'icon.svg', 'icon-192.png', 'icon-512.png', 'icon.svg?v=2', 'icon-192.png?v=2', 'icon-512.png?v=2'];
// 'reload' = never take old copies from the browser cache when installing a new version
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' }))))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== 'mymoney-share').map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Something shared to MyMoney from another app (Android): a picture goes to Credit cards, text to Add expense
  if (e.request.method === 'POST' && url.origin === location.origin && url.pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      const base = new URL('./', self.registration.scope).href;
      try {
        const f = await e.request.formData(), img = f.getAll('image').find(x => x && x.size);
        if (img) {
          const c = await caches.open('mymoney-share');
          await c.put('shared-image', new Response(img, { headers: { 'Content-Type': img.type || 'image/png' } }));
          return Response.redirect(base + 'index.html?shared=image#ccards', 303);
        }
        const q = new URLSearchParams(); ['title', 'text', 'url'].forEach(k => { const v = f.get(k); if (v) q.set(k, v); });
        return Response.redirect(base + 'index.html?' + q.toString(), 303);
      } catch (err) { return Response.redirect(base + 'index.html', 303); }
    })());
    return;
  }
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;   // GitHub API goes straight to the network
  // network first so updates you upload show up; fall back to the cached copy when offline
  // on a slow connection, wait at most 3 seconds and then open the saved copy (the new one is still saved for next time)
  const net = fetch(e.request, { cache: 'no-cache' }).then(r => {   // always ask GitHub whether the file changed
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); } return r;
  });
  const saved = () => caches.match(e.request).then(r => r || caches.match('index.html'));
  e.respondWith(new Promise(done => {
    let settled = false; const finish = r => { if (!settled && r) { settled = true; done(r); } };
    const t = setTimeout(() => saved().then(r => r && finish(r)), 3000);
    net.then(r => { clearTimeout(t); finish(r); }).catch(() => { clearTimeout(t); saved().then(r => finish(r || Response.error())); });
  }));
});

/* Phone notifications sent by the GitHub Action in the data repo */
self.addEventListener('push', e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = { body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'MyMoney', {
    body: d.body || '', icon: 'icon-192.png', badge: 'icon-192.png', tag: d.tag || 'mymoney', renotify: true,
    data: { url: d.url || './' } }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || './', self.registration.scope).href;
  e.waitUntil(clients.matchAll({ type: 'window', includeUncontrolled: true }).then(ws => {
    for (const w of ws) if (w.url.startsWith(self.registration.scope)) { w.navigate(url).catch(() => {}); return w.focus(); }
    return clients.openWindow(url);
  }));
});
