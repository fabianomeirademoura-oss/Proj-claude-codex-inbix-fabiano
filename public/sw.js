// Service worker do painel. Guarda só o que não tem dado de negócio: o CSS, os ícones e a página
// "sem conexão". Telas e CSV passam sempre pela rede (estão atrás do login e mudam com as planilhas).
const VERSAO = 'horizonte-v1';
const ESTATICOS = ['/estilo.css', '/offline.html', '/manifest.webmanifest', '/icones/icone-192.png', '/icones/favicone-32.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSAO).then((c) => c.addAll(ESTATICOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((nomes) => Promise.all(nomes.filter((n) => n !== VERSAO).map((n) => caches.delete(n))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  const caminho = new URL(req.url).pathname;
  if (req.mode === 'navigate') {
    // Sempre a rede; sem conexão, a página de aviso.
    e.respondWith(fetch(req).catch(() => caches.match('/offline.html')));
    return;
  }
  if (caminho === '/estilo.css' || caminho.startsWith('/icones/')) {
    // Rede primeiro (uma mudança de visual aparece na hora); sem conexão, a cópia guardada.
    e.respondWith(caches.open(VERSAO).then((c) => fetch(req)
      .then((r) => { if (r.ok) c.put(req, r.clone()); return r; })
      .catch(() => c.match(req))));
  }
});
