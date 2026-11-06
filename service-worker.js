/* Service worker minimale: serve solo a rendere l'app installabile e a farla
   riaprire (con un messaggio, non con i dati) quando il dispositivo è offline.
   Non tocca mai le chiamate al backend (sono POST, qui gestiamo solo GET) né
   le richieste verso altri domini (Google Fonts, Apps Script, ecc.). */
const CACHE = 'gestionale-impianti-v2';
const SHELL = ['./', './index.html', './manifest.json'];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;
  e.respondWith(
    /* cache: 'no-store' evita che il browser risponda con una copia HTTP già in cache invece di
       andare davvero in rete: senza, dopo un aggiornamento dell'app un utente poteva continuare
       a vedere la versione vecchia per giorni anche ricaricando la pagina. */
    fetch(e.request, { cache: 'no-store' })
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request).then((r) => r || caches.match('./index.html')))
  );
});
