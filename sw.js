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
 *    이 중 URL에 버전이 정확히 고정된 자산(FontAwesome, Google Fonts, 버전 고정된
 *    jsdelivr 패키지)은 "캐시 우선"으로, 나머지(페이지 자체, 버전 미고정인
 *    Tailwind Play CDN)는 "네트워크 우선"으로 서로 다르게 처리한다 - 아래
 *    isVersionPinnedCacheFirstAsset_ 참고.
 * 2) 화면이 열리고 나면, index.html의 기존 localStorage 폴백
 *    (cacheMainDataSnapshot_ / mainDataCache)이 "마지막으로 동기화된 데이터"를
 *    보여주므로, 완전 오프라인이어도 최소한 지난 데이터는 확인할 수 있다.
 *
 * 구글 시트로의 실제 저장/조회(POST 요청)는 이 서비스 워커가 절대 가로채거나
 * 캐싱하지 않는다 - 항상 실제 네트워크로만 나간다.
 * ===================================================================
 */

const CACHE_VERSION = 'dashboard-cache-v4';
// 🩹 [정밀검사 개선] "네트워크 우선, 실패 시 캐시" 전략은 성공한 GET 응답을 전부
// 캐시에 계속 쌓기만 하고 지우는 로직이 없었다. CDN 자산(Tailwind/Chart.js/FontAwesome/
// Google Fonts 등)의 URL이 버전업 등으로 바뀌면 예전 항목이 캐시에 계속 누적될 수 있으므로,
// 캐시 항목 수가 이 값을 넘으면 가장 오래된 것부터 정리한다.
const MAX_CACHE_ENTRIES = 60;

async function trimCache_(cacheName, maxEntries) {
  try {
    const cache = await caches.open(cacheName);
    const keys = await cache.keys();
    if (keys.length <= maxEntries) return;
    // caches.keys()는 저장된 순서(대략 오래된 순)를 반환하므로 앞에서부터 지운다.
    const excess = keys.length - maxEntries;
    for (let i = 0; i < excess; i++) {
      await cache.delete(keys[i]);
    }
  } catch (e) { /* 캐시 정리 실패는 무시 - 기능에는 영향 없음 */ }
}

const CORE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
  './favicon.png'
];

// 🩹 [정밀검사 개선] cache.addAll()은 원자적(all-or-nothing) 동작이라, CORE_ASSETS
// 7개 중 단 하나(배포 직후 CDN/캐시 지연으로 인한 일시적 404 등)만 실패해도 나머지
// 전부가 캐싱되지 않고 통째로 실패했다. 실제로는 "일부만 빠지는" 게 아니라
// "이번 설치에서 코어 자산이 하나도 캐싱 안 됨"이 되어, 그 직후 완전 오프라인으로
// 앱을 열면(비행기모드 등) 흰 화면이 뜰 위험이 있었다. 각 자산을 개별로 fetch+put해서
// 하나가 실패해도 나머지는 정상적으로 캐싱되도록 바꾼다.
// 🩹 [정밀검사 개선] URL 자체에 버전이 고정된 CDN 자산(FontAwesome처럼 경로에 정확한
// 버전 번호가 박혀있는 것, 또는 jsdelivr에서 @1.2.3 형태로 버전을 명시한 것)은 그 URL의
// 내용이 앞으로도 절대 바뀌지 않는다는 게 사실상 보장된다. 이런 자산만 골라서 "캐시 우선"
// 전략을 적용하면, 매번 온라인이어도 CDN 왕복 없이 즉시 렌더링되면서도 기기마다 다른
// 버전이 섞이는 위험은 없다. 반대로 cdn.tailwindcss.com(Play CDN, 버전 미지정 롤링 빌드)이나
// 버전 미지정 jsdelivr 요청은 절대 여기 포함시키지 않는다 - 그 경우엔 "최신"이라는 게 매번
// 달라질 수 있어서 캐시 우선으로 하면 기기 간 드리프트가 생길 수 있다.
const VERSION_PINNED_CACHE_FIRST_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdnjs.cloudflare.com'];

function isVersionPinnedCacheFirstAsset_(url) {
  if (VERSION_PINNED_CACHE_FIRST_HOSTS.includes(url.hostname)) return true;
  // jsdelivr은 경로에 "@정확한버전" 형태(예: /npm/chart.js@4.5.1/...)가 있을 때만
  // 버전 고정으로 간주한다. "@latest"나 버전 생략, 메이저만 지정("@4")은 여전히
  // 시간이 지나면 내용이 바뀔 수 있으므로 캐시 우선에서 제외한다.
  if (url.hostname === 'cdn.jsdelivr.net' && /@\d+\.\d+\.\d+/.test(url.pathname)) return true;
  return false;
}

async function cacheCoreAssetsIndividually_(cache, assets) {
  await Promise.all(assets.map(async (assetUrl) => {
    try {
      const res = await fetch(assetUrl, { cache: 'no-cache' });
      if (res && res.ok) {
        await cache.put(assetUrl, res);
      } else {
        console.warn(`[sw] 코어 자산 캐싱 실패(응답 상태 이상, 무시하고 계속 진행): ${assetUrl}`);
      }
    } catch (err) {
      console.warn(`[sw] 코어 자산 캐싱 실패(무시하고 계속 진행): ${assetUrl}`, err);
    }
  }));
}

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cacheCoreAssetsIndividually_(cache, CORE_ASSETS))
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

  // 버전이 URL에 고정된 CDN 자산(FontAwesome/Google Fonts/버전 고정 jsdelivr 패키지):
  // "캐시 우선, 있으면 즉시 응답 + 뒤에서 조용히 최신본으로 갱신" 전략. 이런 자산은
  // 내용이 바뀔 일이 없으므로 매번 온라인이어도 CDN 왕복을 기다리지 않고 바로 그려진다.
  if (isVersionPinnedCacheFirstAsset_(url)) {
    event.respondWith(
      caches.match(req).then((cached) => {
        const networkFetch = fetch(req)
          .then((res) => {
            if (res && res.ok) {
              const resClone = res.clone();
              caches.open(CACHE_VERSION)
                .then((cache) => cache.put(req, resClone))
                .then(() => trimCache_(CACHE_VERSION, MAX_CACHE_ENTRIES))
                .catch(() => {});
            }
            return res;
          })
          .catch(() => cached || Promise.reject('오프라인이며 캐시도 없음'));
        return cached || networkFetch;
      })
    );
    return;
  }

  // 같은 출처(페이지 자체) 또는 버전 미고정 CDN 자산(Tailwind Play CDN 등):
  // "네트워크 우선, 실패 시 캐시" 전략. 최신 배포본을 최대한 우선 사용하되, 오프라인이면
  // 마지막으로 성공한 캐시로 대체한다.
  // 같은 출처(index.html 등)는 브라우저 HTTP 캐시를 거치지 않고 항상 서버에 재검증해서,
  // 새로 배포한 index.html이 옛 버전에 가려지지 않도록 한다.
  const freshReq = url.origin === self.location.origin ? new Request(req, { cache: 'no-cache' }) : req;
  event.respondWith(
    fetch(freshReq)
      .then((res) => {
        if (res && res.ok) {
          const resClone = res.clone();
          caches.open(CACHE_VERSION)
            .then((cache) => cache.put(req, resClone))
            .then(() => trimCache_(CACHE_VERSION, MAX_CACHE_ENTRIES))
            .catch(() => {});
        }
        return res;
      })
      .catch(() => caches.match(req).then((cached) => cached || (req.mode === 'navigate' ? caches.match('./index.html') : Promise.reject('오프라인이며 캐시도 없음'))))
  );
});
