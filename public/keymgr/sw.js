// Key Manager bản iPhone: lưu sẵn vỏ ứng dụng để mở được khi không có mạng (không lưu dữ liệu key ở đây).
const VERSION = 'km-1.0.0';
const SHELL = ['./', './index.html', './styles.css', './lic.js', './km.js', './app.js', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k.startsWith('km-') && k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
// mạng trước (luôn lấy bản mới khi có mạng), mất mạng thì dùng bản đã lưu
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req).then((r) => { if (r.ok) { const cp = r.clone(); caches.open(VERSION).then((c) => c.put(req.mode === 'navigate' ? './index.html' : req, cp)); } return r; })
    .catch(() => caches.match(req.mode === 'navigate' ? './index.html' : req, { ignoreSearch: true })));
});
