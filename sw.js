/**
 * ===================================================================
 * 서비스 워커 - 오프라인에서도 "앱 껍데기"가 열리도록 캐싱
 * ===================================================================
 * 이 대시보드는 구글 시트(Apps Script)와 실시간으로 통신해야 진짜 데이터를
 * 보여줄 수 있으므로, 서비스 워커가 "네트워크 없이도 최신 데이터"까지
 * 만들어주진 못한다. 대신 이 서비스 워커가 맡는 역할은 두 가지:
 *
 * 1) 페이지 자체(HTML/manifest/아이콘)와 Tailwind/Chart.js/FontAwesome/구글폰트
 *    같은 CDN 리소스를 캐싱해서, 네트워크가 아예 없어도(비행기모드 등) 앱 화면
 *    자체는 열리게 한다("홈 화면에 추가"로 설치했을 때 흰 화면 대신 UI가 보임).
 * 2) 화면이 열리고 나면, index.html의 기존 localStorage 폴백
 *    (cacheMainDataSnapshot_ / mainDataCache)이 "마지막으로 동기화된 데이터"를
 *    보여주므로, 완전 오프라인이어도 최소한 지난 데이터는 확인할 수 있다.
 *
 * 구글 시트로의 실제 저장/조회(POST 요청)는 이 서비스 워커가 절대 가로채거나
 * 캐싱하지 않는다 - 항상 실제 네트워크로만 나간다.
 * ===================================================================
 */

const CACHE_VERSION = 'dashboard-cache-v1';
const CORE_ASSETS = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
  './favicon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(CORE_ASSETS))
      .catch((err) => console.warn('[sw] 코어 자산 캐싱 중 일부 실패(무시 가능):', err))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // GET이 아닌 요청(대부분 구글 시트로의 저장/조회 POST)은 절대 가로채지 않고
  // 그대로 네트워크로 흘려보낸다. respondWith를 호출하지 않으면 브라우저 기본
  // 동작(정상적으로 네트워크 요청)이 그대로 유지된다.
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Apps Script API 도메인(google.com 계열)은 캐싱 대상이 아니다. 지금은 대시보드가
  // 조회도 전부 POST로만 보내므로 사실상 여기 걸릴 일이 없지만, 혹시 모를 GET 요청도
  // 안전하게 네트워크로만 보낸다(캐시에 절대 남기지 않음).
  if (url.hostname.endsWith('google.com') || url.hostname.endsWith('googleusercontent.com')) {
    return;
  }

  // 같은 출처(페이지 자체) 또는 CDN 정적 자산: "네트워크 우선, 실패 시 캐시" 전략.
  // 최신 배포본을 최대한 우선 사용하되, 오프라인이면 마지막으로 성공한 캐시로 대체한다.
  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res && res.ok) {
          const resClone = res.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put(req, resClone)).catch(() => {});
        }
        return res;
      })
      .catch(() => caches.match(req).then((cached) => cached || Promise.reject('오프라인이며 캐시도 없음')))
  );
});
