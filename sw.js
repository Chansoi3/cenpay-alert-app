// Service Worker สำหรับ Cenpay Alert PWA
// ทำให้ติดตั้งเป็นแอปได้ + โหลดเร็วขึ้นผ่าน cache

const CACHE_NAME = 'cenpay-alert-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './assets/logo.png',
  './assets/favicon-32.png',
  './assets/icon-192.png',
  './assets/icon-512.png',
  // CDN assets (cache ได้แต่อาจมี CORS)
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js',
  'https://unpkg.com/vue@2.6.14/dist/vue.js',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
  'https://fonts.googleapis.com/css2?family=Mali:wght@400;500;600;700&family=Sarabun:wght@400;500;600;700&display=swap'
];

// ====== INSTALL: cache ไฟล์หลัก ======
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        // cache ทีละไฟล์ ถ้าบางไฟล์ล้มเหลว ไม่ให้ทั้ง install พัง
        return Promise.allSettled(
          ASSETS.map(url => cache.add(url).catch(err => console.log('Cache miss:', url, err.message)))
        );
      })
      .then(() => self.skipWaiting())
  );
});

// ====== ACTIVATE: ลบ cache เก่า ======
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      );
    }).then(() => self.clients.claim())
  );
});

// ====== FETCH: network-first สำหรับคำขอใหม่, fallback เป็น cache ======
self.addEventListener('fetch', (event) => {
  // ข้ามคำขอที่ไม่ใช่ GET (เช่น POST ไป Apps Script)
  if (event.request.method !== 'GET') return;

  // ข้าม request ไป Apps Script (ต้องการข้อมูลสดเสมอ)
  if (event.request.url.includes('script.google.com')) return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // ถ้าสำเร็จ → cache สำเนา แล้วคืน response
        if (response && response.status === 200 && response.type === 'basic') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // ถ้า offline → ใช้ cache
        return caches.match(event.request).then(cached => cached || caches.match('./index.html'));
      })
  );
});
