/**
 * ===================================================================
 * 종목별 7차 분할매수 데이터 - 구글시트 연동 API (Apps Script)
 * ===================================================================
 *
 * [설치 방법]
 * 1) 구글시트 상단 메뉴에서 [확장 프로그램] > [Apps Script] 클릭
 * 2) 기존 코드를 전부 지우고 이 파일의 내용을 전부 붙여넣기
 * 3) 저장 (Ctrl+S 또는 플로피 아이콘)
 * 4) 우측 상단 [배포] > [배포 관리] > 연필 아이콘 > 버전을 "새 버전"으로 선택 후 배포
 * 5) [중요] 자동 백업이 실제로 동작하게 하려면, 이 편집기 상단 함수 목록에서
 *    "createDailyBackupTrigger"를 선택한 뒤 ▶ 실행 버튼을 한 번 눌러줘야 한다.
 *    이 단계를 건너뛰면 runDailyBackup_()/restoreFromBackupFile_() 등 백업/복원
 *    로직 자체는 있어도 매일 자동으로 실행되는 트리거가 걸리지 않아서, 나중에
 *    실수로 데이터를 잘못 편집했을 때 되돌릴 백업이 하나도 없을 수 있다.
 *    (실행 후 권한 승인 팝업이 뜨면 허용해주면 된다. 자세한 내용은 아래 "자동 백업" 섹션 참고)
 * 6) [중요] 인증 토큰 설정: 아래 함수 목록에서 "setApiToken_"을 선택 → 코드 안의
 *    함수 정의를 잠깐 보고 원하는 토큰 문자열을 정한 뒤, 편집기 상단의 "실행" 버튼
 *    옆 드롭다운에서 함수를 고르고 인자를 넣어 1회 실행(또는 아래 setApiToken_ 함수
 *    설명 참고). 이제 토큰은 이 코드 파일이 아니라 구글의 PropertiesService(스크립트
 *    속성)에 저장되므로, 이 파일을 복사/백업/공유해도 토큰이 함께 노출되지 않는다.
 *    실행 후 대시보드(index.html)의 "Apps Script 동기화 설정" 카드에도 동일한 토큰을
 *    입력해야 한다. (자세한 내용은 아래 "🔐 인증" 섹션 참고)
 * ===================================================================
 */

const SHEET_NAME = '분할매수데이터';

// doPost가 스크립트 락을 잡고 있는 동안 true. (실행마다 전역 변수가 새로 초기화된다)
let SCRIPT_LOCK_HELD_ = false;

// 일일 기록(수익금/계좌평가액/종목합산평가금액/종목별 매입금액)이 저장되는 메인 시트 이름과 열 위치
const MAIN_SHEET_NAME = '시트1';
const MAIN_COLS = {
  date: 1,         // A: 날짜
  revenue: 2,      // B: 수익금
  cumRevenue: 3,   // C: 수익금 총합 (참고용, 자동 계산)
  accountVal: 4,   // D: 전체계좌
  mdd: 5,          // E: 최고점 대비 하락률 (참고용, 자동 계산)
  dailyReturn: 6,  // F: 일별 수익률 (참고용, 자동 계산)
  memo: 7,         // G: 특이사항
  stockSumAgg: 8,  // H: 종목 합산 평가금액
  tickerStart: 9   // I부터: 종목별 매입금액
};

// 종목 데이터 뒤(맨 끝)에 자동으로 붙는 "현금" 컬럼의 헤더명
const CASH_HEADER = '현금';

// 🩹 [수정 4] 며칠 기록을 건너뛰고 나중에 입력할 때, saveDailyRecord_가 그 사이 날짜를
// 자동으로 채워 넣는 행의 "특이사항" 칸에 남기는 표시. 프런트는 이 문구를 보고
// 실제로 입력하지 않은 "미기록일"을 일평균 계산의 분모에서 제외한다.
const AUTO_FILL_MEMO_MARKER = '자동채움(미기록일)';

// ===================================================================
// 🔐 인증
// ===================================================================
// ⚠️ 배포 전 반드시 토큰을 아무도 짐작할 수 없는 임의의 문자열로 설정하세요.
//    (예: 32자 이상의 랜덤 문자열. 대시보드(index.html)의
//    "Apps Script 동기화 설정" 카드의 "인증 토큰" 입력란에 동일한 값을 입력해야 합니다.)
//    설정하지 않고 그대로 배포하면 이 웹앱은 인증 없이 열려있는 것과 같습니다.
//
// 🔒 [정밀검사 개선] 예전에는 토큰을 이 코드 안에 상수(API_TOKEN)로 평문 그대로 적어뒀다.
// 이 파일을 다른 곳에 복사/백업하거나(예: 깃허브에 실수로 올리는 등) 화면 공유를 하면
// 토큰이 그대로 노출된다. PropertiesService(스크립트 속성)에 토큰을 저장해두고 코드에는
// 남기지 않는 방식으로 바꾼다.
//
// [설정 방법] 이 편집기 상단 함수 목록에서 "setApiToken_"을 선택 → 아래 안내대로 인자를
// 채워 ▶ 실행 (또는 함수 실행 대화상자에서 파라미터를 직접 입력해도 됨). 한 번만 하면 된다.
const API_TOKEN_PROP_KEY_ = 'apiToken_v1';
// 하위 호환용 폴백: 스크립트 속성에 토큰이 없을 때만 사용되며, 이 값 그대로면 항상 인증 실패
// 처리된다(= 기본값 그대로 배포해도 안전). 예전처럼 이 상수를 직접 바꿔 써도 계속 동작한다.
const API_TOKEN_FALLBACK_ = 'CHANGE_ME_TO_YOUR_OWN_SECRET_TOKEN';

// 최초 1회 실행: 토큰을 코드가 아닌 PropertiesService에 저장한다. 실행 방법:
// 1) 위 함수 목록에서 "setApiToken_" 선택 → ▶ 실행하면 파라미터 입력창이 뜨지 않는
//    구버전 UI도 있으므로, 대신 이 편집기 하단 "실행 로그" 옆의 "▶ 실행" 드롭다운에서
//    "함수 실행..."을 골라 인자를 넣거나, 아래 코드에서 'CHANGE_ME...' 부분을 원하는 토큰
//    문자열로 잠깐 바꿔 실행 후 원래대로 되돌려도 된다. 실행 후에는 이 값이 스크립트
//    속성에 저장되므로 코드를 다시 되돌려도 토큰은 유지된다.
function setApiToken_(token) {
  token = String(token || '').trim();
  if (!token || token === API_TOKEN_FALLBACK_) {
    throw new Error('유효한 토큰 문자열을 인자로 넣어 실행하세요. 예: setApiToken_("아무도-못-맞추는-32자-이상-무작위-문자열")');
  }
  PropertiesService.getScriptProperties().setProperty(API_TOKEN_PROP_KEY_, token);
  Logger.log('토큰이 스크립트 속성에 저장되었습니다. 이제 대시보드의 "인증 토큰" 입력란에도 동일한 값을 입력하세요.');
}

// 저장된 토큰을 지우고 다시 코드 상수(API_TOKEN_FALLBACK_, 기본값이면 항상 인증 실패)로
// 되돌린다. 토큰 유출이 의심될 때 즉시 이 웹앱 접근을 막는 용도로도 쓸 수 있다.
function clearApiToken_() {
  PropertiesService.getScriptProperties().deleteProperty(API_TOKEN_PROP_KEY_);
  Logger.log('저장된 토큰을 삭제했습니다. setApiToken_()으로 새 토큰을 다시 설정하기 전까지 이 웹앱은 모든 요청을 거부합니다.');
}

function getApiToken_() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty(API_TOKEN_PROP_KEY_);
    if (stored) return stored;
  } catch (e) { /* PropertiesService 오류 시 아래 하위 호환 상수로 폴백 */ }
  return API_TOKEN_FALLBACK_;
}

// 🩹 [정밀검사 개선] 단순 문자열 === 비교는 이론적으로 타이밍 공격(비교가 얼마나
// 빨리 "다르다"고 판정되는지를 측정해 정답 토큰을 한 글자씩 추측하는 공격)에
// 노출된다. 개인용 대시보드 수준에서 실질적 위험은 낮지만, 길이와 무관하게 항상
// 전체를 끝까지 비교하는 상수 시간(constant-time) 비교로 바꿔 그 여지를 없앤다.
// Apps Script(V8)에는 Node의 crypto.timingSafeEqual 같은 내장 함수가 없어 직접
// 구현한다: 두 문자열을 UTF-8 바이트로 바꾼 뒤, 길이가 달라도 항상 "두 길이 중 더
// 긴 쪽" 만큼 순회하며 모든 바이트를 XOR로 누적한다(어느 한 바이트가 다른 순간
// 바로 return하지 않음 → 첫 글자가 틀리든 마지막 글자가 틀리든 걸리는 시간이 같다).
function timingSafeStringEquals_(a, b) {
  const bytesA = Utilities.newBlob(String(a || '')).getBytes();
  const bytesB = Utilities.newBlob(String(b || '')).getBytes();
  const maxLen = Math.max(bytesA.length, bytesB.length);
  let diff = bytesA.length === bytesB.length ? 0 : 1;
  for (let i = 0; i < maxLen; i++) {
    const byteA = i < bytesA.length ? bytesA[i] : 0;
    const byteB = i < bytesB.length ? bytesB[i] : 0;
    diff |= (byteA ^ byteB);
  }
  return diff === 0;
}

function checkToken_(token) {
  const apiToken = getApiToken_();
  return !!apiToken
    && apiToken !== API_TOKEN_FALLBACK_
    && timingSafeStringEquals_(token, apiToken);
}

// 🆕 [정밀검사 개선] 서버 측 입력 검증
// 지금까지는 날짜/숫자 형식 검증을 전적으로 프런트(parseStrictNumber_ 등)에 의존했다.
// 토큰만 알면 누구든 이 웹앱에 직접 POST를 보낼 수 있으므로, 잘못된 형식의 날짜가
// 들어와 시트 정렬(locateDateRow_의 이진 탐색 전제)이 깨지는 사고를 막기 위해
// 서버에서도 최소한의 형식 검증을 한 번 더 거친다.
function isValidDateStr_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return false;
  const [yy, mm, dd] = String(s).split('-').map(Number);
  const d = new Date(yy, mm - 1, dd);
  return d.getFullYear() === yy && d.getMonth() === mm - 1 && d.getDate() === dd;
}

function isValidNumberInput_(v) {
  if (v === undefined || v === null || v === '') return true; // 빈 값은 각 함수에서 0 등으로 기본 처리
  return !isNaN(Number(v));
}

// ===================================================================
// 🆕 [버그 수정] 쓰기 요청 재시도 시 중복 처리 방지 (idempotency)
// ===================================================================
// 대시보드(index.html)의 fetchWithRetry_는 네트워크 타임아웃/일시 오류 시 같은 요청
// 본문을 그대로 최대 2번 더 재전송한다. saveStep_/addTicker_/saveDailyRecord_ 등 대부분의
// 쓰기 함수는 "이미 있으면 갱신, 없으면 추가"(upsert) 방식이라 같은 요청이 두 번 들어와도
// 결과가 같지만, addBuyMgmtRecord_(확정 수익 기록 추가)처럼 호출할 때마다 새 ID를 만들어
// 무조건 새 행을 추가하는 함수는 그렇지 않다. Apps Script 콜드스타트 등으로 서버 응답이
// 늦어져 클라이언트가 타임아웃으로 판단해 재전송하면, 실제로는 이미 처리가 끝난 요청이
// 한 번 더 실행되어 같은 확정 수익금이 중복 저장될 수 있었다.
//
// 이를 막기 위해 프런트가 요청마다 함께 보내는 clientRequestId(요청 1건당 한 번만 생성되며,
// 재시도에도 동일 값이 유지됨 - withToken_ 참고)를 기준으로, 이미 처리해서 응답을 만들어둔
// 요청이면 그 액션을 다시 실행하지 않고 저장해둔 응답을 그대로 돌려준다. 이 캐시는 짧은
// 시간(REQUEST_DEDUPE_TTL_SEC_)만 유지하면 충분하다 - 재시도는 보통 몇 초 안에 끝난다.
const REQUEST_DEDUPE_TTL_SEC_ = 120;

function requestDedupeCacheKey_(clientRequestId) {
  return 'reqdedupe_' + String(clientRequestId);
}

function getDedupedResponse_(clientRequestId) {
  if (!clientRequestId) return null;
  try {
    const raw = CacheService.getScriptCache().get(requestDedupeCacheKey_(clientRequestId));
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function storeDedupedResponse_(clientRequestId, responsePayload) {
  if (!clientRequestId) return;
  try {
    const serialized = JSON.stringify(responsePayload);
    // CacheService는 값 하나당 약 100KB 제한이 있다. 이를 넘으면(예: 매우 큰 grid 응답)
    // 캐싱을 건너뛴다 - 중복 방지는 못 받지만 방금 실행한 저장 자체는 이미 끝난 상태이므로
    // 기능이 깨지지는 않는다.
    if (serialized.length < 90000) {
      CacheService.getScriptCache().put(requestDedupeCacheKey_(clientRequestId), serialized, REQUEST_DEDUPE_TTL_SEC_);
    }
  } catch (e) { /* 캐시 저장 실패는 무시 */ }
}

// ===================================================================
// 🆕 [정밀검사 개선 - 문제 B] 화면 설정(평일만 토글 / 차트별 "기준일" 캘린더 필터)
// PC/모바일 동기화
// ===================================================================
// 지금까지 이 값들은 대시보드(index.html)의 localStorage에만 저장되어 있어서, 한
// 기기에서 바꿔도 다른 기기에는 전혀 반영되지 않고 각자 다른 값을 계속 기억하는
// 문제가 있었다. 이 값들을 시트 자체에 컬럼을 추가하는 대신 Apps Script의
// PropertiesService(스크립트 속성)에 작은 JSON 하나로 저장해두면, 어느 기기에서
// 조회하든 같은 값을 받아본다.
const UI_SETTINGS_PROP_KEY_ = 'uiSettings_v1';
// 차트별 "기준일" 필터가 적용되는 8개 화면(일평균 KPI 카드 + 상세 테이블 + 콤보/누적/
// MDD/연도별/월별수익률/종목별월별 차트).
const UI_SETTINGS_CHART_KEYS_ = ['combo', 'acc', 'mdd', 'yearly', 'monthlyReturn', 'tickerMonthly', 'table', 'avgRevenue'];

// 🆕 [기능 추가] 도넛 차트 "목표 비중(%)" - 종목/현금 이름을 키로 하는 임의 개수의
// 맵이라 UI_SETTINGS_CHART_KEYS_처럼 고정된 키 목록으로 다루지 않고, 저장된 항목을
// 그대로 신뢰하되 형식이 이상한 항목만 걸러낸다. 키 길이를 제한해 실수로 아주 긴
// 문자열이 들어와 저장 용량을 낭비하는 것을 막는다.
const TARGET_ALLOCATION_KEY_MAX_LEN_ = 30;

// 🆕 [정밀검사 개선 - 설계상 한계 해소] 목표 계좌 평가액 / 목표 달성 알림 켬·끔 / 표 밀도는
// 지금까지 대시보드(index.html)의 localStorage에만 저장되는 "이 기기 전용" 설정이었다.
// PC와 폰에서 서로 다른 값을 보게 되는 것 자체는 의도된 설계였지만, 지난번 "평일만" 토글/
// 차트별 기준일 필터/목표 비중(%)처럼 여러 기기에서 같은 값을 보고 싶을 수도 있으므로,
// 나머지 화면 설정과 동일하게 이 값들도 PropertiesService에 함께 저장해 기기 간 동기화한다.
// localStorage는 그대로 남겨두어 오프라인일 때의 폴백 캐시 역할을 계속 한다.
function getUiSettings_() {
  let parsed = {};
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(UI_SETTINGS_PROP_KEY_);
    parsed = raw ? JSON.parse(raw) : {};
  } catch (e) { parsed = {}; }

  const chartStart = {};
  UI_SETTINGS_CHART_KEYS_.forEach(key => {
    const v = parsed.chartStart && parsed.chartStart[key];
    chartStart[key] = (typeof v === 'string' && isValidDateStr_(v)) ? v : '';
  });

  const targetAllocations = {};
  if (parsed.targetAllocations && typeof parsed.targetAllocations === 'object') {
    Object.keys(parsed.targetAllocations).forEach(k => {
      const key = String(k || '').slice(0, TARGET_ALLOCATION_KEY_MAX_LEN_);
      if (!key) return;
      const v = Number(parsed.targetAllocations[k]);
      if (!isNaN(v) && v >= 0) targetAllocations[key] = v;
    });
  }

  const goalAccountValNum = Number(parsed.goalAccountVal);
  const goalAccountVal = (!isNaN(goalAccountValNum) && goalAccountValNum > 0) ? goalAccountValNum : null;

  return {
    weekdayOnly: !!parsed.weekdayOnly,
    chartStart: chartStart,
    targetAllocations: targetAllocations,
    goalAccountVal: goalAccountVal,
    goalNotifyEnabled: !!parsed.goalNotifyEnabled,
    tableDensityCompact: !!parsed.tableDensityCompact
  };
}

// patch로 넘어온 필드만 반영하고(부분 갱신), 나머지는 기존 저장값을 그대로 유지한다.
// 날짜 형식이 아닌 값(빈 문자열 제외)은 조용히 무시해서 잘못된 값으로 다른 기기의
// 화면이 깨지는 일을 막는다.
function saveUiSettings_(patch) {
  const current = getUiSettings_();
  patch = patch || {};

  if (typeof patch.weekdayOnly === 'boolean') current.weekdayOnly = patch.weekdayOnly;

  if (patch.chartStart && typeof patch.chartStart === 'object') {
    UI_SETTINGS_CHART_KEYS_.forEach(key => {
      if (!Object.prototype.hasOwnProperty.call(patch.chartStart, key)) return;
      const v = patch.chartStart[key];
      if (v === '' || v === null || v === undefined) {
        current.chartStart[key] = '';
      } else if (isValidDateStr_(v)) {
        current.chartStart[key] = v;
      }
      // 그 외 형식이 이상한 값은 무시(기존 값 유지)
    });
  }

  // 🆕 [기능 추가] 목표 비중(%) 기기 간 동기화. patch.targetAllocations는
  // { 종목명: 숫자(설정) | ''(그 항목 삭제) } 형태의 부분 갱신 맵으로, 값이 ''/null/undefined면
  // 그 종목의 목표 비중을 완전히 삭제하고, 그 외에는 0 이상의 숫자만 반영한다.
  if (patch.targetAllocations && typeof patch.targetAllocations === 'object') {
    if (!current.targetAllocations || typeof current.targetAllocations !== 'object') current.targetAllocations = {};
    Object.keys(patch.targetAllocations).forEach(rawKey => {
      const key = String(rawKey || '').slice(0, TARGET_ALLOCATION_KEY_MAX_LEN_);
      if (!key) return;
      const v = patch.targetAllocations[rawKey];
      if (v === '' || v === null || v === undefined) {
        delete current.targetAllocations[key];
      } else {
        const num = Number(v);
        if (!isNaN(num) && num >= 0) current.targetAllocations[key] = num;
      }
    });
  }

  // 🆕 [정밀검사 개선 - 설계상 한계 해소] 목표 계좌 평가액. ''/null/undefined면 목표를
  // 완전히 삭제(더 이상 목표를 두지 않음)한 것으로 간주하고, 그 외에는 0보다 큰 숫자만 반영한다.
  if (Object.prototype.hasOwnProperty.call(patch, 'goalAccountVal')) {
    const v = patch.goalAccountVal;
    if (v === '' || v === null || v === undefined) {
      current.goalAccountVal = null;
    } else {
      const num = Number(v);
      if (!isNaN(num) && num > 0) current.goalAccountVal = num;
    }
  }

  // 🆕 [정밀검사 개선 - 설계상 한계 해소] 목표 달성 브라우저 알림 켬·끔. 알림 권한(Notification
  // API) 자체는 브라우저별로 별도 승인이 필요해 기기 간 동기화가 불가능하지만, "켜고 싶다는
  // 의도"만큼은 다른 화면 설정과 동일하게 저장해두면, 다른 기기에서도 (그 기기에서 권한을
  // 승인한 뒤) 다시 켜지 않아도 되도록 돕는다.
  if (typeof patch.goalNotifyEnabled === 'boolean') current.goalNotifyEnabled = patch.goalNotifyEnabled;

  // 🆕 [정밀검사 개선 - 설계상 한계 해소] 상세 표 밀도(좁게/기본) 토글.
  if (typeof patch.tableDensityCompact === 'boolean') current.tableDensityCompact = patch.tableDensityCompact;

  PropertiesService.getScriptProperties().setProperty(UI_SETTINGS_PROP_KEY_, JSON.stringify(current));
  return current;
}

// ------------------- 진입점 -------------------

/**
 * 🔒 [보안 강화] GET 경로 완전 차단.
 *
 * 예전에는 doGet도 doPost와 동일한 action들을 지원해서(하위 호환 목적),
 * 브라우저 주소창에 "https://script.google.com/.../exec?action=getAll&token=..."
 * 형태로 직접 호출할 수 있었다. 이 방식은 인증 토큰이 그대로 URL에 담기기 때문에
 * 브라우저 히스토리 / 프록시·회사 네트워크 접근 로그 / 공유된 스크린샷 등을 통해
 * 토큰이 노출될 위험이 있다. 대시보드(index.html)는 이미 오래 전부터 모든 조회/저장을
 * POST 본문으로만 보내도록 바뀌었으므로, GET 경로는 더 이상 필요하지 않다.
 *
 * 지금부터 doGet은 어떤 action/token을 붙여 호출해도 데이터를 절대 반환하지 않고,
 * "POST로 호출하세요"라는 안내만 돌려준다. 기존에 저장해 둔 배포 URL이 있어도
 * (예: 즐겨찾기에 GET 링크를 저장해뒀던 경우) 더 이상 데이터가 새어나가지 않는다.
 */
function doGet(e) {
  return jsonOut_({
    success: false,
    error: '이 API는 GET 요청을 지원하지 않습니다. 반드시 POST 요청으로 호출해주세요(대시보드는 이미 POST만 사용합니다).'
  });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (parseErr) {
    return jsonOut_({ success: false, error: '요청 본문(JSON) 파싱 실패: ' + String(parseErr) });
  }

  if (!checkToken_(body.token)) {
    return jsonOut_({ success: false, error: '인증 실패: token이 올바르지 않습니다.' });
  }

  // 🔐 [정밀검사 개선] 조회(read) 액션은 시트를 변경하지 않으므로 쓰기 락(lock)이 필요 없다.
  // 예전에는 대시보드가 이 조회들을 doGet(GET 쿼리스트링, ?action=...&token=...)으로
  // 호출했는데, 그러면 인증 토큰이 브라우저 주소창/히스토리/네트워크 로그에 그대로
  // 노출될 수 있었다. 프런트엔드가 동일한 조회를 POST 본문으로도 보낼 수 있게 여기서
  // 먼저 처리하고, 락 없이 즉시 반환한다(저장 중이어도 조회가 불필요하게 대기하지 않음).
  const READ_ACTIONS_ = { getAll: 1, getGrid: 1, getMainData: 1, getBuyMgmtRecords: 1, getBootstrap: 1, getUiSettings: 1, listBackups: 1, exportSnapshot: 1 };
  if (READ_ACTIONS_[body.action]) {
    try {
      return handleReadAction_(body);
    } catch (err) {
      return jsonOut_({ success: false, error: String(err) });
    }
  }

  // 🆕 [버그 수정] 이 요청(clientRequestId)을 이미 처리해서 응답까지 만들어둔 적이 있다면,
  // 락을 잡거나 액션을 다시 실행하지 않고 그때 응답을 그대로 돌려준다. fetchWithRetry_의
  // 자동 재시도로 같은 addBuyMgmtRecord 요청이 두 번 실행되어 확정 수익금이 중복 저장되는
  // 사고를 여기서 원천 차단한다(다른 쓰기 액션들에도 동일하게 적용되어 앞으로 비슷한 재시도
  // 중복 문제가 생기는 것도 함께 막는다).
  const dedupedResponse = getDedupedResponse_(body.clientRequestId);
  if (dedupedResponse) return jsonOut_(dedupedResponse);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 직렬화 (중복 행 생성 방지)
  } catch (lockErr) {
    return jsonOut_({ success: false, error: '서버가 바빠서 저장을 못 했어요. 잠시 후 다시 시도해주세요.' });
  }
  // 🩹 [정밀검사 개선] 이 실행이 스크립트 락을 이미 잡고 있음을 알린다. runDailyBackup_가
  // 락 없이 도는 예약 트리거 경로에서만 스스로 락을 잡도록 구분하기 위한 표시다.
  SCRIPT_LOCK_HELD_ = true;
  // 락 획득 후 재확인: 첫 요청이 아직 실행 중일 때 온 재시도가 중복 실행되는 것을 막는다.
  const dedupedAfterLock = getDedupedResponse_(body.clientRequestId);
  if (dedupedAfterLock) { SCRIPT_LOCK_HELD_ = false; lock.releaseLock(); return jsonOut_(dedupedAfterLock); }

  // 🆕 [버그 수정] 예전에는 각 case가 곧바로 jsonOut_(...)을 return해서, "응답을 만든 뒤
  // clientRequestId로 캐싱해둔다"는 공통 처리를 걸 지점이 없었다. 이제 각 case는
  // responsePayload에 결과 객체만 대입하고 break하며, switch가 끝난 뒤 딱 한 곳에서
  // 캐싱 + 응답 반환을 처리한다(로직 자체는 예전과 동일, 응답을 만드는 시점만 한 곳으로 모음).
  let responsePayload;
  try {
    switch (body.action) {
      case 'saveStep': {
        const sheet = getOrCreateSheet_();
        if (!isValidNumberInput_(body.alloc) || !isValidNumberInput_(body.bought)) { responsePayload = { success: false, error: '할당/매수 금액은 숫자여야 합니다.' }; break; }
        saveStep_(sheet, body.ticker, body.step, body.alloc, body.bought);
        responsePayload = { success: true, ...readAll_(sheet) };
        break;
      }
      case 'addTicker': {
        const sheet = getOrCreateSheet_();
        addTicker_(sheet, body.ticker);
        responsePayload = { success: true, ...readAll_(sheet) };
        break;
      }
      case 'deleteTicker': {
        const sheet = getOrCreateSheet_();
        deleteTicker_(sheet, body.ticker);
        responsePayload = { success: true, ...readAll_(sheet) };
        break;
      }
      case 'saveDailyRecord': {
        if (!isValidDateStr_(body.date)) {
          responsePayload = { success: false, error: '날짜 형식이 올바르지 않습니다(YYYY-MM-DD): ' + body.date };
          break;
        }
        if (!isValidNumberInput_(body.revenue) || !isValidNumberInput_(body.accountVal) || !isValidNumberInput_(body.stockSumVal)) {
          responsePayload = { success: false, error: '수익금/계좌평가액/종목합산평가금액은 숫자여야 합니다.' };
          break;
        }
        // 🆕 [정밀검사 개선] 지금까지 revenue/accountVal/stockSumVal만 서버에서 숫자 형식을
        // 검증했고, tickerValues(종목별 매입금액)는 saveDailyRecord_ 안에서 검증 없이
        // Number(v) || 0으로 처리되어, 값이 깨져 들어오면 에러 없이 조용히 0으로 기록될 수
        // 있었다(예: 특정 종목의 실제 보유금액이 실수로 사라짐). 다른 필드들과 동일한 기준으로
        // 여기서 한 번 더 형식을 확인해, 잘못된 값이면 저장 자체를 막고 어떤 종목이
        // 문제인지 알려준다.
        if (body.tickerValues && typeof body.tickerValues === 'object') {
          const invalidTicker = Object.keys(body.tickerValues)
            .find(t => !isValidNumberInput_(body.tickerValues[t]));
          if (invalidTicker) {
            responsePayload = {
              success: false,
              error: `'${invalidTicker}' 종목의 값이 숫자가 아닙니다: ${body.tickerValues[invalidTicker]}`
            };
            break;
          }
        }
        const mainSheet = getMainSheet_();
        const guardErr = validateRecordDateRange_(mainSheet, body.date);
        if (guardErr) { responsePayload = { success: false, error: guardErr }; break; }
        const result = saveDailyRecord_(mainSheet, body);
        responsePayload = { success: true, cash: result.cash, stockSum: result.stockSum, recalcWarning: result.recalcWarning || null };
        break;
      }
      case 'deleteDailyRecord': {
        // 상세 데이터 테이블에서 특정 날짜의 기록 행을 완전히 삭제한다.
        // 삭제 후에는 그 날짜 위치부터 누적수익금/MDD/일별수익률을 다시 계산해야
        // 이후 날짜들의 파생값이 예전 값 그대로 남지 않는다.
        if (!body.date) { responsePayload = { success: false, error: 'date 누락' }; break; }
        if (!isValidDateStr_(body.date)) {
          responsePayload = { success: false, error: '날짜 형식이 올바르지 않습니다(YYYY-MM-DD): ' + body.date };
          break;
        }
        const mainSheet = getMainSheet_();
        const deleteResult = deleteDailyRecord_(mainSheet, body.date);
        responsePayload = { success: true, ...readMainData_(mainSheet), recalcWarning: (deleteResult && deleteResult.recalcWarning) || null };
        break;
      }
      case 'deleteTickerFromMain': {
        // 종목을 완전히 삭제할 때 메인 시트(시트1)의 해당 종목 컬럼도 통째로 지운다.
        // (예전 zeroTickerColumn 액션은 마지막 행만 0으로 바꾸고 컬럼 자체는 남겨둬서,
        //  삭제한 종목의 빈 컬럼이 시트에 계속 쌓이는 문제가 있었음)
        const mainSheet = getMainSheet_();
        deleteTickerColumnFromMain_(mainSheet, body.ticker);
        responsePayload = { success: true };
        break;
      }
      case 'saveGridCell': {
        if (!body.sheetName) { responsePayload = { success: false, error: 'sheetName 누락' }; break; }
        const gridSheet = getOrCreateGridSheet_(body.sheetName);
        saveGridCell_(gridSheet, body.rowLabel, body.ticker, body.value);
        responsePayload = { success: true, ...readGrid_(gridSheet) };
        break;
      }
      case 'saveGridRow': {
        // 한 행의 여러 칸(예: 일평균수익금/월평균수익금/월평균수익률/연환산수익률)을
        // 요청 1건으로 저장한다. saveGridCell을 필드 개수만큼 동시에 여러 번 호출하면
        // 락(lock) 경쟁으로 일부가 무작위로 실패할 수 있어, 프런트에서 값들을 모아
        // 한 번의 doPost 호출로 순차 저장하도록 통합한 액션.
        if (!body.sheetName) { responsePayload = { success: false, error: 'sheetName 누락' }; break; }
        const gridSheet = getOrCreateGridSheet_(body.sheetName);
        saveGridRow_(gridSheet, body.rowLabel, body.values || {});
        responsePayload = { success: true, ...readGrid_(gridSheet) };
        break;
      }
      case 'deleteGridTicker': {
        if (!body.sheetName) { responsePayload = { success: false, error: 'sheetName 누락' }; break; }
        const gridSheet = getOrCreateGridSheet_(body.sheetName);
        deleteGridTickerColumn_(gridSheet, body.ticker);
        responsePayload = { success: true, ...readGrid_(gridSheet) };
        break;
      }
      case 'addGridTicker': {
        if (!body.sheetName) { responsePayload = { success: false, error: 'sheetName 누락' }; break; }
        const gridSheet = getOrCreateGridSheet_(body.sheetName);
        addGridTickerColumn_(gridSheet, body.ticker);
        responsePayload = { success: true, ...readGrid_(gridSheet) };
        break;
      }
      case 'addGridRow': {
        if (!body.sheetName) { responsePayload = { success: false, error: 'sheetName 누락' }; break; }
        const gridSheet = getOrCreateGridSheet_(body.sheetName);
        addGridRow_(gridSheet, body.rowLabel);
        responsePayload = { success: true, ...readGrid_(gridSheet) };
        break;
      }
      case 'deleteGridRow': {
        // 차수 기록 등에서 특정 행(예: "5차수") 전체를 완전히 삭제한다.
        // 그 행에 저장돼 있던 모든 종목의 날짜/수량·메모/수익금 칸이 함께 삭제된다.
        if (!body.sheetName) { responsePayload = { success: false, error: 'sheetName 누락' }; break; }
        const gridSheet = getOrCreateGridSheet_(body.sheetName);
        deleteGridRow_(gridSheet, body.rowLabel);
        responsePayload = { success: true, ...readGrid_(gridSheet) };
        break;
      }
      case 'addBuyMgmtRecord': {
        // '종목별 매수 관리'에서 (계좌 × 종목) 칸의 💰 토글을 열어 '기록 추가'를 눌렀을 때:
        // 그 시점의 날짜/종목/계좌/메모/수익금을 별도의 "확정 내역" 시트에 영구 저장한다.
        // 이렇게 저장된 값은 원본 할당/매수 칸과 무관하게 그대로 남아 "당일 수익금" 합산에 반영된다.
        // 🩹 [버그 수정] 이 액션은 호출할 때마다 새 ID로 무조건 새 행을 추가하므로(=재시도에
        // 안전하지 않으므로), 이 doPost 진입점 상단의 clientRequestId 중복 차단이 특히 중요하다.
        if (!isValidDateStr_(body.date)) {
          responsePayload = { success: false, error: '날짜 형식이 올바르지 않습니다(YYYY-MM-DD): ' + body.date };
          break;
        }
        if (!isValidNumberInput_(body.profit)) {
          responsePayload = { success: false, error: '수익금은 숫자여야 합니다.' };
          break;
        }
        const recSheet = getOrCreateBuyMgmtRecordsSheet_();
        const id = addBuyMgmtRecord_(recSheet, {
          date: body.date, ticker: body.ticker, account: body.account, memo: body.memo, profit: body.profit
        });
        responsePayload = { success: true, id: id, ...readBuyMgmtRecords_(recSheet) };
        break;
      }
      case 'updateBuyMgmtRecord': {
        if (!body.id) { responsePayload = { success: false, error: 'id 누락' }; break; }
        if (!isValidDateStr_(body.date)) {
          responsePayload = { success: false, error: '날짜 형식이 올바르지 않습니다(YYYY-MM-DD): ' + body.date };
          break;
        }
        if (!isValidNumberInput_(body.profit)) {
          responsePayload = { success: false, error: '수익금은 숫자여야 합니다.' };
          break;
        }
        const recSheet = getOrCreateBuyMgmtRecordsSheet_();
        updateBuyMgmtRecord_(recSheet, body.id, {
          date: body.date, ticker: body.ticker, account: body.account, memo: body.memo, profit: body.profit
        });
        responsePayload = { success: true, ...readBuyMgmtRecords_(recSheet) };
        break;
      }
      case 'deleteBuyMgmtRecord': {
        if (!body.id) { responsePayload = { success: false, error: 'id 누락' }; break; }
        const recSheet = getOrCreateBuyMgmtRecordsSheet_();
        deleteBuyMgmtRecord_(recSheet, body.id);
        responsePayload = { success: true, ...readBuyMgmtRecords_(recSheet) };
        break;
      }
      case 'addCashflow': {
        if (!isValidDateStr_(body.date) || body.amount === '' || body.amount == null || isNaN(Number(body.amount)) || Number(body.amount) === 0) {
          responsePayload = { success: false, error: '날짜/금액(0 제외, 입금 +, 출금 -)을 확인해주세요.' }; break;
        }
        const cfSheet = getOrCreateCashflowSheet_();
        const cfId = addCashflow_(cfSheet, body);
        responsePayload = { success: true, id: cfId, cashflows: readCashflows_(cfSheet) };
        break;
      }
      case 'deleteCashflow': {
        if (!body.id) { responsePayload = { success: false, error: 'id 누락' }; break; }
        const cfSheet = getOrCreateCashflowSheet_();
        deleteCashflow_(cfSheet, body.id);
        responsePayload = { success: true, cashflows: readCashflows_(cfSheet) };
        break;
      }
      case 'restoreBackup': {
        // 💡 [기능 추가] 지금까지는 restoreFromBackupFile_를 Apps Script 편집기에서
        // fileId를 손으로 바꿔가며 수동 실행해야 했다. 대시보드의 "백업에서 복원" 모달에서
        // 바로 호출할 수 있게 doPost 액션으로 노출한다. 복원은 시트 내용을 통째로
        // 덮어쓰는 되돌릴 수 없는 작업이므로, 다른 저장 액션들과 동일하게 락을 잡은
        // 상태에서 실행하고(동시에 다른 저장과 겹치지 않도록), 복원 직전 현재 상태도
        // 스냅샷으로 한 번 남겨서 "복원의 복원"이 가능하게 한다.
        // 🩹 [정밀검사 개선 - 문제 1] snapshotBeforeDestructiveAction_는 하루 1회 제한이
        // 있어, 같은 날 먼저 다른 삭제로 이미 스냅샷을 찍어뒀다면 여기서 다시 호출해도
        // 아무 백업이 생기지 않아 "복원 직전 상태는 자동으로 백업된다"는 확인창 안내가
        // 지켜지지 않을 수 있었다. 복원 전용으로 하루 제한 없이 항상 백업하는
        // snapshotBeforeRestore_를 대신 사용한다.
        if (!body.fileId) { responsePayload = { success: false, error: 'fileId 누락' }; break; }
        if (!listBackupFiles_().some(f => f.id === body.fileId)) { responsePayload = { success: false, error: '백업 폴더에 없는 파일입니다.' }; break; }
        const preRestore = snapshotBeforeRestore_();
        if (!preRestore.ok) {
          responsePayload = { success: false, error: '복원 직전 상태를 백업하지 못해 복원을 중단했습니다(시트는 변경되지 않았습니다): ' + preRestore.error };
          break;
        }
        const restoreResult = restoreFromBackupFile_(body.fileId);
        responsePayload = { success: true, restoredSheets: restoreResult.restoredSheets, backedUpAt: restoreResult.backedUpAt };
        break;
      }
      case 'importSnapshot': {
        // 🆕 [기능 추가 - JSON 백업 불러오기] 대시보드에서 사용자가 직접 고른 JSON
        // 백업 파일(exportSnapshot으로 받아둔 무손실 스냅샷)을 다시 시트로 복원한다.
        // fileId 기반의 기존 restoreBackup(드라이브 백업)과 달리, 요청 본문에 스냅샷
        // 전체(body.snapshot)가 직접 실려온다.
        if (!body.snapshot || typeof body.snapshot !== 'object') {
          responsePayload = { success: false, error: 'snapshot 데이터가 없습니다.' };
          break;
        }
        // 시트를 건드리기 전에 먼저 구조를 검증한다. 여기서 실패하면 스냅샷 백업(드라이브 쓰기)
        // 조차 만들지 않고 바로 반환한다.
        const preCheckErr = validateSnapshotStructure_(body.snapshot);
        if (preCheckErr) { responsePayload = { success: false, error: preCheckErr }; break; }

        const preImport = snapshotBeforeRestore_();
        if (!preImport.ok) {
          responsePayload = { success: false, error: '복원 직전 상태를 백업하지 못해 복원을 중단했습니다(시트는 변경되지 않았습니다): ' + preImport.error };
          break;
        }
        const importResult = applySnapshot_(body.snapshot);
        if (!importResult.success) { responsePayload = importResult; break; }
        responsePayload = {
          success: true,
          restoredSheets: importResult.restoredSheets,
          skipped: importResult.skipped,
          formulasRestored: importResult.formulasRestored,
          formulasSkipped: importResult.formulasSkipped,
          backedUpAt: importResult.backedUpAt
        };
        break;
      }
      case 'saveUiSettings': {
        // 🆕 [정밀검사 개선 - 문제 B] "평일만" 토글 / 차트별 "기준일" 필터를 PropertiesService에
        // 저장해 PC/모바일이 같은 값을 공유하게 한다. 시트를 변경하지 않지만, 다른 쓰기
        // 액션들과 같은 락 경로를 타도 되는 아주 가벼운 작업이라 여기 switch에 그대로 둔다.
        const updated = saveUiSettings_(body.uiSettings || {});
        responsePayload = { success: true, uiSettings: updated };
        break;
      }
      default:
        responsePayload = { success: false, error: 'unknown action: ' + body.action };
    }
  } catch (err) {
    responsePayload = { success: false, error: String(err) };
  } finally {
    storeDedupedResponse_(body.clientRequestId, responsePayload); // 락 해제 전에 저장
    SCRIPT_LOCK_HELD_ = false;
    lock.releaseLock();
    // 방금 시트에 어떤 형태로든 쓰기 시도가 있었으므로, 다음 getBootstrap 호출이
    // 오래된(stale) 캐시를 돌려주지 않도록 항상 무효화한다(성공/실패와 무관하게).
    invalidateBootstrapCache_();
  }

  // 🆕 [버그 수정] 방금 만든 응답을 clientRequestId로 캐싱해둔다. 이 요청이 (타임아웃 등으로)
  // 재전송되어 doPost가 다시 호출되면, 함수 맨 위의 getDedupedResponse_가 이 값을 찾아내
  // 액션을 다시 실행하지 않고 바로 돌려준다.
  storeDedupedResponse_(body.clientRequestId, responsePayload);
  return jsonOut_(responsePayload);
}

/**
 * 🔐 [정밀검사 개선] doPost로 들어온 조회(read) 액션 전용 핸들러.
 * token이 URL 쿼리스트링이 아니라 POST 본문에 담겨 오는 것을 제외하면, 예전에
 * doGet이 처리하던 것과 동일한 action들을 여기서 처리한다.
 * (🔒 [보안 강화] doGet은 이제 어떤 action도 처리하지 않고 항상 에러만 반환하므로,
 *  조회는 오직 이 경로 하나로만 들어온다.)
 */
function handleReadAction_(body) {
  switch (body.action) {
    case 'getAll': {
      const sheet = getOrCreateSheet_();
      return jsonOut_({ success: true, ...readAll_(sheet) });
    }
    case 'getGrid': {
      if (!body.sheetName) {
        return jsonOut_({ success: false, error: 'sheetName 파라미터가 누락되었습니다.' });
      }
      const gridSheet = getOrCreateGridSheet_(body.sheetName);
      return jsonOut_({ success: true, ...readGrid_(gridSheet) });
    }
    case 'getMainData': {
      const mainSheet = getMainSheet_();
      return jsonOut_({ success: true, ...readMainData_(mainSheet) });
    }
    case 'getBuyMgmtRecords': {
      const recSheet = getOrCreateBuyMgmtRecordsSheet_();
      return jsonOut_({ success: true, ...readBuyMgmtRecords_(recSheet) });
    }
    case 'getBootstrap': {
      return jsonOut_({ success: true, ...buildBootstrap_() });
    }
    case 'getUiSettings': {
      return jsonOut_({ success: true, uiSettings: getUiSettings_() });
    }
    case 'exportSnapshot': {
      // 🆕 [기능 추가 - JSON 백업 내보내기] 대시보드의 "JSON 백업 다운로드" 버튼이
      // 화면에서 파싱한 값이 아니라, 시트 원본 그대로의 무손실 스냅샷을 받아갈 수
      // 있게 한다. 드라이브에 저장하지 않고 응답으로만 내려준다는 점만 runDailyBackup_와 다르다.
      return jsonOut_({ success: true, ...buildSnapshot_(), restorableSheets: Object.keys(getImportSnapshotAllowedSheets_()) });
    }
    case 'listBackups': {
      // 💡 [기능 추가] 대시보드에서 "백업 목록 보기 → 복원" 버튼을 쓸 수 있도록,
      // 지금까지 Apps Script 편집기에서 listBackupFiles_()를 수동 실행해야 확인할 수
      // 있던 백업 목록을 API로도 조회할 수 있게 한다. Date 객체는 JSON으로 직접
      // 직렬화되지 않으므로 ISO 문자열로 바꿔서 내려준다.
      const backups = listBackupFiles_().map(f => ({ id: f.id, name: f.name, createdAt: f.createdAt.toISOString() }));
      return jsonOut_({ success: true, backups: backups });
    }
    default:
      return jsonOut_({ success: false, error: 'unknown read action: ' + body.action });
  }
}

// ===================================================================
// 🆕 [기능 7] 초기 로딩 일괄 조회 (getBootstrap)
// ===================================================================
// 대시보드 최초 로딩 시 순서대로 8번 호출하던
// (분할매수데이터 → 메인시트 → 매수목록/차수기록/확정내역/메모장/메모장1/수익률데이터)
// 요청을 한 번의 POST 호출(action=getBootstrap)로 묶어서 반환한다. Apps Script 웹앱은
// 요청마다 콜드스타트 지연(1~3초)이 있을 수 있어, 특히 모바일 첫 로딩 체감 속도가 크게 개선된다.
const BOOTSTRAP_GRID_SHEETS = {
  buyList: '매수목록',
  stepLog: '차수기록',
  notepad: '메모장',
  notepad1: '메모장1',
  yieldData: '수익률데이터'
};

// 🆕 [정밀검사 개선] getBootstrap 결과를 짧은 시간(20초) 동안 캐싱한다.
// Apps Script 웹앱은 콜드스타트 시 요청당 1~3초가 걸릴 수 있는데, 실제로는 같은 화면을
// 짧은 간격으로 여러 번 새로고침(당겨서 새로고침, 탭 전환 등)하는 경우가 잦다. 그 사이에는
// 시트를 다시 읽지 않고 캐시를 즉시 반환해 체감 속도를 올린다. 쓰기 액션(saveXxx/addXxx/
// deleteXxx 등)이 하나라도 실행되면 doPost의 finally에서 이 캐시를 즉시 지우므로, 저장 직후
// 새로고침해도 오래된(stale) 데이터를 보는 일은 없다.
// 🩹 [정밀검사 개선] 예전에는 캐시 키 하나를 put/remove 했다. 그래서 (1) 읽는 도중 쓰기가 끝나
// invalidate된 뒤에 "오래된" 결과가 다시 put되어 20초간 남을 수 있었고, (2) 직렬화 결과가
// 90KB를 넘으면 캐시가 아예 꺼졌다. 이제는 "세대(generation) 번호"를 두어, 읽기 시작 때의
// 세대와 저장 직전의 세대가 같을 때만 캐시에 쓰고, 큰 결과는 여러 조각(chunk)으로 나눠 저장한다.
const BOOTSTRAP_CACHE_GEN_KEY_ = 'bootstrap_gen';
const BOOTSTRAP_CACHE_TTL_SEC_ = 20;
const BOOTSTRAP_CACHE_CHUNK_CHARS_ = 30000;  // 한글 3바이트 기준으로도 100KB 제한 안에 들어오게
const BOOTSTRAP_CACHE_MAX_CHUNKS_ = 30;

function newBootstrapGen_() {
  return String(Date.now()) + '_' + Math.floor(Math.random() * 1000000);
}

function getBootstrapGen_() {
  const cache = CacheService.getScriptCache();
  let gen = cache.get(BOOTSTRAP_CACHE_GEN_KEY_);
  if (!gen) {
    gen = newBootstrapGen_();
    cache.put(BOOTSTRAP_CACHE_GEN_KEY_, gen, 21600);
  }
  return gen;
}

function invalidateBootstrapCache_() {
  try { CacheService.getScriptCache().put(BOOTSTRAP_CACHE_GEN_KEY_, newBootstrapGen_(), 21600); } catch (e) { /* 캐시 서비스 오류는 무시 */ }
}

function readBootstrapCache_(gen) {
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get('bootstrap_n_' + gen));
  if (!count || count < 1 || count > BOOTSTRAP_CACHE_MAX_CHUNKS_) return null;
  const keys = [];
  for (let i = 0; i < count; i++) keys.push('bootstrap_c_' + gen + '_' + i);
  const got = cache.getAll(keys);
  let joined = '';
  for (let i = 0; i < count; i++) {
    const part = got[keys[i]];
    if (part === undefined || part === null) return null; // 일부 조각이 만료/유실 → 캐시 미사용
    joined += part;
  }
  return JSON.parse(joined);
}

function writeBootstrapCache_(gen, serialized) {
  const chunks = [];
  for (let i = 0; i < serialized.length; i += BOOTSTRAP_CACHE_CHUNK_CHARS_) {
    chunks.push(serialized.substring(i, i + BOOTSTRAP_CACHE_CHUNK_CHARS_));
  }
  if (chunks.length === 0 || chunks.length > BOOTSTRAP_CACHE_MAX_CHUNKS_) return; // 너무 크면 캐시 생략
  const obj = {};
  chunks.forEach((c, i) => { obj['bootstrap_c_' + gen + '_' + i] = c; });
  obj['bootstrap_n_' + gen] = String(chunks.length);
  CacheService.getScriptCache().putAll(obj, BOOTSTRAP_CACHE_TTL_SEC_);
}

// 🩹 [정밀검사 개선] 읽기 경로(락 없음)가 시트를 자동 생성하면, 쓰기 요청과 겹칠 때 같은 이름의
// 시트가 중복 생성되는 경합이 생길 수 있다. 하나라도 없을 때만 락을 잡고 만든다.
function ensureBootstrapSheetsExist_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = [SHEET_NAME, BUYMGMT_RECORDS_SHEET, CASHFLOW_SHEET] // 시트1은 없으면 첫 시트로 대체되므로 제외
    .concat(Object.keys(BOOTSTRAP_GRID_SHEETS).map(k => BOOTSTRAP_GRID_SHEETS[k]));
  if (names.every(n => ss.getSheetByName(n))) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 못 잡으면 예외 → 조회 실패로 응답(클라이언트가 다음 주기에 재시도)
  try {
    getOrCreateSheet_();
    getMainSheet_();
    Object.keys(BOOTSTRAP_GRID_SHEETS).forEach(k => getOrCreateGridSheet_(BOOTSTRAP_GRID_SHEETS[k]));
    getOrCreateBuyMgmtRecordsSheet_();
    getOrCreateCashflowSheet_();
  } finally {
    lock.releaseLock();
  }
}

function buildBootstrap_() {
  let gen = null;
  try {
    gen = getBootstrapGen_();
    const cached = readBootstrapCache_(gen);
    if (cached) return cached;
  } catch (e) { /* 캐시 조회 실패 시 그냥 새로 계산 */ }

  ensureBootstrapSheetsExist_();

  const splitBuySheet = getOrCreateSheet_();
  const splitBuy = readAll_(splitBuySheet);

  const mainSheet = getMainSheet_();
  const mainData = readMainData_(mainSheet);

  const grids = {};
  Object.keys(BOOTSTRAP_GRID_SHEETS).forEach(key => {
    const gridSheet = getOrCreateGridSheet_(BOOTSTRAP_GRID_SHEETS[key]);
    grids[key] = readGrid_(gridSheet);
  });

  const recSheet = getOrCreateBuyMgmtRecordsSheet_();
  const buyMgmtRecords = readBuyMgmtRecords_(recSheet);

  const tz = Session.getScriptTimeZone();
  const result = {
    tickers: splitBuy.tickers,
    data: splitBuy.data,
    matrix: mainData.matrix,
    grids: grids,
    records: buyMgmtRecords.records,
    uiSettings: getUiSettings_(),
    cashflows: readCashflows_(getOrCreateCashflowSheet_()),
    backupStatus: getLastBackupInfo_(),
    // 🩹 [정밀검사 개선] 서버(스크립트) 시간대. 화면 시간대와 어긋나면 클라이언트가 경고한다.
    serverTimeZone: tz,
    serverUtcOffset: Utilities.formatDate(new Date(), tz, 'Z')
  };

  try {
    // 읽는 동안 쓰기가 끝났다면(세대 번호가 바뀜) 오래된 결과이므로 캐시하지 않는다.
    if (gen !== null && getBootstrapGen_() === gen) {
      writeBootstrapCache_(gen, JSON.stringify(result));
    }
  } catch (e) { /* 캐시 저장 실패는 무시 - 기능에는 영향 없음 */ }

  return result;
}

// ------------------- 시트 준비 -------------------

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['티커', '차수', '할당($)', '매수($)']]);
    sheet.getRange(1, 6, 1, 1).setValue('종목목록');
    sheet.setColumnWidths(1, 4, 90);
    sheet.setColumnWidth(6, 100);
  }
  return sheet;
}

// ------------------- 조회 -------------------

function readAll_(sheet) {
  const lastRow = sheet.getLastRow();
  const data = {};

  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    rows.forEach(r => {
      const ticker = String(r[0] || '').trim().toUpperCase();
      const step = Number(r[1]);
      if (!ticker || !step) return;
      if (!data[ticker]) data[ticker] = {};
      data[ticker][step] = {
        alloc: r[2] === '' ? '' : r[2],
        bought: r[3] === '' ? '' : r[3]
      };
    });
  }

  const tickers = [];
  if (lastRow >= 2) {
    const tCol = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    tCol.forEach(r => {
      const t = String(r[0] || '').trim().toUpperCase();
      if (t) tickers.push(t);
    });
  }

  return { tickers: tickers, data: data };
}

// ------------------- 수정 -------------------

function findRow_(sheet, ticker, step) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === ticker && Number(rows[i][1]) === step) {
      return i + 2;
    }
  }
  return -1;
}

function saveStep_(sheet, ticker, step, alloc, bought) {
  ticker = String(ticker).trim().toUpperCase();
  step = Number(step);
  const rowIdx = findRow_(sheet, ticker, step);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 3, 1, 2).setValues([[alloc, bought]]);
  } else {
    // 🩹 [정밀검사 개선] appendRow로 직접 쓰면 ticker가 "005930" 같은 순수 숫자로만
    // 이루어진 한국 종목코드일 경우, 구글시트가 이를 숫자(5930)로 자동 변환해 앞자리 0이
    // 사라질 수 있다. 다른 그리드 시트들과 동일하게, 쓰기 전에 A열을 텍스트 서식으로
    // 먼저 고정한다.
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 1).setNumberFormat('@');
    sheet.getRange(newRow, 1, 1, 4).setValues([[ticker, step, alloc, bought]]);
  }
}

function addTicker_(sheet, ticker) {
  ticker = String(ticker).trim().toUpperCase();
  if (!ticker) return;

  const lastRow = sheet.getLastRow();
  let existing = [];
  if (lastRow >= 2) {
    existing = sheet.getRange(2, 6, lastRow - 1, 1).getValues()
      .map(r => String(r[0]).trim().toUpperCase())
      .filter(Boolean);
  }
  if (existing.includes(ticker)) return;

  let lastF = 1;
  if (lastRow >= 2) {
    const fCol = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    for (let i = fCol.length - 1; i >= 0; i--) { if (String(fCol[i][0]).trim()) { lastF = i + 2; break; } }
  }
  const nextRow = lastF + 1;
  // 🩹 [정밀검사 개선] 종목목록(F열)도 saveStep_의 A열과 동일한 순수 숫자 티커
  // 자동 변환 위험이 있어, 쓰기 전에 텍스트 서식을 먼저 적용한다.
  sheet.getRange(nextRow, 6, 1, 1).setNumberFormat('@').setValue(ticker);
}

function deleteTicker_(sheet, ticker) {
  ticker = String(ticker).trim().toUpperCase();

  let lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]).trim().toUpperCase() === ticker) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const tRange = sheet.getRange(2, 6, lastRow - 1, 1);
    const values = tRange.getValues().map(r => String(r[0]).trim().toUpperCase());
    const filtered = values.filter(t => t && t !== ticker);
    tRange.clearContent();
    if (filtered.length > 0) {
      sheet.getRange(2, 6, filtered.length, 1).setValues(filtered.map(t => [t]));
    }
  }
}

// ===================================================================
// 📝 일일 기록 (수익금 / 계좌평가액 / 종목별 평가금) - 메인 시트 저장
// ===================================================================

function getMainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  return sheet;
}

// 🩹 [정밀검사 개선] 중복 종목 헤더 방어. 시트를 직접 편집해 같은 종목명 헤더가
// 두 열 이상 생기면, 예전에는 뒤 열이 앞 열을 조용히 덮어써서(map[티커] = 열번호)
// 앞선 열의 값이 조회에서 아무 경고 없이 무시됐다(값 손실은 아니고 "안 보이게"
// 되는 정도). 지난번 "중복 날짜" 방어와 동일한 방식으로, 여기서는 최소한 실행
// 로그에 경고를 남겨 문제를 알아챌 수 있게 하고(동작 자체는 하위 호환을 위해
// "나중 열 우선"을 유지), 근본적인 정리는 cleanupDuplicateTickerHeaders_()로 하도록
// 안내한다.
function getTickerColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const map = {};
  const seenCols = {};
  if (lastCol >= MAIN_COLS.tickerStart) {
    const headers = sheet.getRange(1, MAIN_COLS.tickerStart, 1, lastCol - MAIN_COLS.tickerStart + 1).getValues()[0];
    headers.forEach((h, idx) => {
      const t = String(h || '').trim().toUpperCase();
      if (!t || t === CASH_HEADER.toUpperCase()) return;
      const col = MAIN_COLS.tickerStart + idx;
      if (map[t] !== undefined) {
        Logger.log(`⚠️ 메인 시트(${sheet.getName()})에 종목 헤더 "${t}"가 ${seenCols[t]}열과 ${col}열에 중복 존재합니다. ` +
          `${seenCols[t]}열의 값은 조회에서 무시되고 ${col}열만 사용됩니다. cleanupDuplicateTickerHeaders_()로 정리할 수 있습니다.`);
      }
      seenCols[t] = col;
      map[t] = col;
    });
  }
  return map;
}

function ensureTickerColumn_(sheet, ticker) {
  ticker = ticker.toUpperCase();
  if (ticker === CASH_HEADER.toUpperCase()) return null;
  const map = getTickerColumns_(sheet);
  if (map[ticker]) return map[ticker];

  // 🩹 [정밀검사 개선] 메인 시트(시트1)의 종목 헤더도 순수 숫자 티커(한국 종목코드 등)일 경우
  // 구글시트가 자동으로 숫자로 변환할 수 있다. getTickerColumns_는 헤더를 String(h)로 읽으므로,
  // 숫자로 변환된 헤더("5930")는 원래 티커("005930")와 더 이상 일치하지 않아 매번 새 컬럼이
  // 생기는 등 데이터가 어긋난다. 헤더 쓰기 전에 항상 텍스트 서식을 먼저 적용해 막는다.
  const cashCol = findCashColumn_(sheet);
  if (cashCol) {
    sheet.insertColumnBefore(cashCol);
    sheet.getRange(1, cashCol).setNumberFormat('@').setValue(ticker);
    return cashCol;
  }

  const newCol = Math.max(sheet.getLastColumn() + 1, MAIN_COLS.tickerStart);
  sheet.getRange(1, newCol).setNumberFormat('@').setValue(ticker);
  return newCol;
}

function findCashColumn_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol >= MAIN_COLS.tickerStart) {
    const headers = sheet.getRange(1, MAIN_COLS.tickerStart, 1, lastCol - MAIN_COLS.tickerStart + 1).getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i] || '').trim().toUpperCase() === CASH_HEADER.toUpperCase()) {
        return MAIN_COLS.tickerStart + i;
      }
    }
  }
  return null;
}

function ensureCashColumn_(sheet) {
  const existing = findCashColumn_(sheet);
  if (existing) return existing;
  const newCol = Math.max(sheet.getLastColumn() + 1, MAIN_COLS.tickerStart);
  sheet.getRange(1, newCol).setValue(CASH_HEADER);
  return newCol;
}

function formatDate_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '').trim();
}

function addDays_(dateStr, n) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function findDateRow_(sheet, dateStr) {
  return locateDateRow_(sheet, dateStr).row;
}

/**
 * 🩹 [정밀검사 개선] 날짜 행을 찾을 때, "찾는 날짜가 없다면 정렬이 깨지지 않게
 * 어느 행 앞에 끼워 넣어야 하는지"(insertBeforeRow)도 함께 계산해서 반환한다.
 * 메인 시트는 항상 날짜 오름차순이라는 전제(자동채움/삽입 로직이 보장) 하에
 * 이진 탐색으로 찾아 비교 횟수를 줄인다.
 *
 * [버그 수정] 예전에는 날짜 행을 못 찾으면 writeRowRaw_가 무조건 시트 맨 끝에
 * 새 행을 추가했다. "마지막 기록보다 미래" 날짜라면 문제없지만, 중간에 비어있는
 * 과거 날짜(진짜 구멍)를 저장하려 하면 맨 끝에 붙어버려 날짜 순서가 깨지고,
 * 이후 recalcDerivedFrom_(누적수익금/MDD/일별수익률)가 잘못된 순서로 계산되는
 * 문제가 있었다. insertBeforeRow를 이용해 정확한 위치에 삽입하도록 고쳤다.
 */
function locateDateRow_(sheet, dateStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { row: -1, insertBeforeRow: 2 };

  const dates = sheet.getRange(2, MAIN_COLS.date, lastRow - 1, 1).getValues()
    .map(r => formatDate_(r[0]));

  let lo = 0, hi = dates.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (dates[mid] === dateStr) return { row: mid + 2, insertBeforeRow: mid + 2 };
    if (dates[mid] < dateStr) lo = mid + 1; else hi = mid - 1;
  }
  // 못 찾았으면 lo는 "이 날짜보다 큰 첫 번째 행"의 배열 인덱스 → 그 위치 앞에 끼워 넣으면
  // 오름차순이 유지된다 (모든 날짜보다 미래면 lo === dates.length, 즉 맨 끝 다음 행).
  return { row: -1, insertBeforeRow: lo + 2 };
}

function getLastDataRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const lastCol = Math.max(sheet.getLastColumn(), MAIN_COLS.tickerStart - 1);
  const values = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
  return { rowIndex: lastRow, values: values, dateStr: formatDate_(values[MAIN_COLS.date - 1]) };
}

function extractTickerValuesFromRow_(rowValues, tickerCols) {
  const result = {};
  Object.keys(tickerCols).forEach(t => {
    const col = tickerCols[t];
    result[t] = Number(rowValues[col - 1]) || 0;
  });
  return result;
}

function deleteTickerColumnFromMain_(sheet, ticker) {
  if (!ticker) return;
  const tickerCols = getTickerColumns_(sheet);
  const col = tickerCols[String(ticker).trim().toUpperCase()];
  if (!col) return;
  snapshotBeforeDestructiveAction_();
  sheet.deleteColumn(col);
}

/**
 * 🆕 readMainData_
 * 메인 시트(시트1) 전체를 헤더 포함 2차원 문자열 배열(matrix)로 반환한다.
 * 프런트엔드가 기존에 구글시트를 "링크가 있는 모든 사용자에게 공개"로 열어두고
 * gviz(/gviz/tq) 엔드포인트로 직접 읽던 것을, 이 인증된 Apps Script 엔드포인트로
 * 대체하기 위한 용도. 시트를 더 이상 공개로 설정하지 않아도 대시보드가 동작한다.
 */
function readMainData_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), MAIN_COLS.tickerStart - 1);
  if (lastRow < 1 || lastCol < 1) return { matrix: [] };

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => (h === null || h === undefined) ? '' : String(h));

  const dataRows = [];

  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    rows.forEach(r => {
      dataRows.push(r.map(v => {
        if (v instanceof Date) {
          return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return (v === null || v === undefined) ? '' : String(v);
      }));
    });
  }

  // 🩹 [정밀검사 개선] 시트가 정상적으로 유지된다면 locateDateRow_의 이진탐색 삽입 덕분에
  // 이미 날짜 오름차순이지만, 사용자가 시트를 직접 열어 행을 드래그하거나 데이터를
  // 복사/붙여넣기해서 순서가 깨질 수 있다. 프런트(parseAndRenderData)의 MDD/일별수익률
  // 계산이 "행 순서 = 날짜 순서"라는 전제에 의존하므로, 여기서 응답 직전에 한 번 더
  // 날짜 문자열 기준으로 명시적으로 정렬해 그 전제를 서버가 보장해준다.
  dataRows.sort((a, b) => {
    const da = String(a[MAIN_COLS.date - 1] || '');
    const db = String(b[MAIN_COLS.date - 1] || '');
    if (da < db) return -1;
    if (da > db) return 1;
    return 0;
  });

  // 🩹 [정밀검사 개선] 중복 날짜 방어. 정상적인 저장 흐름(saveDailyRecord_ → locateDateRow_)은
  // upsert 방식이라 같은 날짜가 두 번 저장될 수 없지만, 사용자가 시트를 직접 열어 행을
  // 복사/붙여넣기하면 같은 날짜가 여러 번 존재할 수 있다. 이를 그대로 내려보내면 클라이언트가
  // 그 날짜의 수익금을 두 번(또는 그 이상) 합산해 누적수익금/MDD/일별수익률이 왜곡된다.
  // 근본적인 정리는 cleanupMainSheetDuplicateDates_()로 시트 자체를 고쳐야 하지만, 그 전까지도
  // 화면이 최소한 왜곡되지 않도록 여기서 한 번 더 방어한다: 같은 날짜의 행이 여럿이면 시트에서
  // 더 나중에 나온(=더 최근에 편집됐을 가능성이 높은) 행만 남긴다.
  const dedupedRows = [];
  const seenDates = {};
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const d = String(dataRows[i][MAIN_COLS.date - 1] || '');
    if (seenDates[d]) continue;
    seenDates[d] = true;
    dedupedRows.push(dataRows[i]);
  }
  dedupedRows.reverse();

  return { matrix: [headerRow, ...dedupedRows] };
}

function saveDailyRecord_(sheet, body) {
  const date = String(body.date).trim();
  const revenue = Number(body.revenue) || 0;
  const accountVal = Number(body.accountVal) || 0;
  const stockSumVal = Number(body.stockSumVal) || 0;
  const memo = body.memo || '';
  const tickerValuesInput = body.tickerValues || {};

  ensureMainHeaderLabels_(sheet);

  Object.keys(tickerValuesInput).forEach(t => ensureTickerColumn_(sheet, t.toUpperCase()));
  const tickerCols = getTickerColumns_(sheet);

  // 이번 호출로 값이 바뀐 가장 앞쪽 행 번호. 파생 컬럼(누적수익금/MDD/일별수익률) 재계산은
  // 이 값이 확정된 뒤 딱 한 번만 호출한다(건너뛴 날짜가 여러 개여도 recalc는 1회).
  let earliestChangedRow = null;

  // ---- 1) 빈 날짜 자동 채우기 (배치 처리) ----
  // 🩹 [정밀검사 개선] 예전에는 공백 날짜마다 writeRowRaw_ → locateDateRow_를 호출했는데,
  // locateDateRow_는 호출될 때마다 시트의 날짜 열 "전체"를 다시 읽어와 이진탐색을 한다.
  // 즉 공백이 N일이고 기존 기록이 M행이면 대략 O(N×M)의 시트 읽기가 발생해, 오랜만에
  // 접속해 몇 달~1년치를 한 번에 채우면 Apps Script 실행시간 제한(6분)에 걸려 저장이
  // 실패할 위험이 있었다. 공백 날짜들은 항상 "시트의 실제 마지막 행 바로 다음"에
  // 순서대로 이어 붙는 것이 보장되므로(last는 getLastDataRow_가 반환한 진짜 마지막 행,
  // gapDates는 전부 last.dateStr보다 미래·date보다 과거), 매번 다시 위치를 찾을 필요 없이
  // writeGapRowsBatch_로 한 번의 getRange/setValues에 몰아서 기록한다.
  const last = getLastDataRow_(sheet);
  if (last && last.dateStr && last.dateStr < date) {
    const lastTickerValues = extractTickerValuesFromRow_(last.values, tickerCols);
    const lastAccountVal = Number(last.values[MAIN_COLS.accountVal - 1]) || 0;
    const lastStockSumAgg = Number(last.values[MAIN_COLS.stockSumAgg - 1]) || 0;

    const gapDates = [];
    let cursor = addDays_(last.dateStr, 1);
    while (cursor < date) {
      gapDates.push(cursor);
      cursor = addDays_(cursor, 1);
    }

    if (gapDates.length > 0) {
      earliestChangedRow = writeGapRowsBatch_(sheet, gapDates, lastAccountVal, lastStockSumAgg, tickerCols, lastTickerValues);
    }
  }

  // ---- 2) 수정: 기존 행이 존재하는 경우 해당 행의 값을 기준(baseline)으로 설정 ----
  const loc = locateDateRow_(sheet, date);
  const existingRow = loc.row;
  let targetRowValues = null;
  if (existingRow > 0) {
    const lastCol = Math.max(sheet.getLastColumn(), MAIN_COLS.tickerStart - 1);
    targetRowValues = sheet.getRange(existingRow, 1, 1, lastCol).getValues()[0];
  } else {
    // 🩹 [정밀검사 개선] 예전에는 새 날짜 행이 없으면 무조건 "시트의 마지막 행"을 기준값으로
    // 삼았다. 보통은(오늘 날짜를 새로 추가하는 일반적인 경우) 마지막 행이 곧 바로 이전
    // 기록이라 문제없지만, date가 과거의 진짜 빈 자리(구멍)라면 마지막 행은 이 날짜보다
    // "미래" 기록이라 기준으로 삼으면 안 된다. 이 날짜가 삽입될 위치 바로 앞 행(=이 날짜
    // 이전 중 가장 가까운 날짜)을 기준으로 삼도록 고쳤다.
    const prevRow = loc.insertBeforeRow - 1;
    if (prevRow >= 2) {
      const lastCol = Math.max(sheet.getLastColumn(), MAIN_COLS.tickerStart - 1);
      targetRowValues = sheet.getRange(prevRow, 1, 1, lastCol).getValues()[0];
    }
  }

  const baseline = targetRowValues ? extractTickerValuesFromRow_(targetRowValues, tickerCols) : {};
  const finalTickerValues = {};
  Object.keys(tickerCols).forEach(t => {
    finalTickerValues[t] = (tickerValuesInput[t] !== undefined && tickerValuesInput[t] !== '')
      ? Number(tickerValuesInput[t]) || 0
      : (baseline[t] || 0);
  });

  const finalRowIdx = writeRowRaw_(sheet, date, revenue, accountVal, memo, stockSumVal, tickerCols, finalTickerValues);
  if (earliestChangedRow === null || finalRowIdx < earliestChangedRow) earliestChangedRow = finalRowIdx;

  const recalcResult = recalcDerivedFrom_(sheet, earliestChangedRow);

  const cash = accountVal - stockSumVal;
  const result = { cash: cash, stockSum: stockSumVal };
  if (recalcResult && recalcResult.rowsRecalculated > RECALC_WARNING_ROW_THRESHOLD) {
    result.recalcWarning = `과거 기록 수정으로 ${recalcResult.rowsRecalculated.toLocaleString()}행을 다시 계산했습니다. 데이터가 계속 쌓이면 이 작업이 느려지거나 실패할 수 있습니다.`;
  }
  return result;
}

function ensureMainHeaderLabels_(sheet) {
  const current = String(sheet.getRange(1, MAIN_COLS.stockSumAgg).getValue() || '').trim();
  if (current !== '종목합산평가금액') {
    sheet.getRange(1, MAIN_COLS.stockSumAgg).setValue('종목합산평가금액');
  }
}

/**
 * ⚡ writeRowRaw_
 * 기본 입력값(날짜/수익금/계좌평가액/메모/종목데이터/현금)만 1번의 API 호출로 기록한다.
 * 파생 컬럼(누적수익금/MDD/일수익률) 재계산은 하지 않고, 실제로 쓴 행 번호만 반환한다.
 *
 * [성능] saveDailyRecord_에서 건너뛴 날짜를 여러 개 자동 채울 때 이 함수를 반복 호출하고,
 * recalcDerivedFrom_는 그 뒤 딱 한 번만(가장 앞쪽으로 바뀐 행부터) 호출하도록 분리했다.
 * 예전에는 writeRow_ 안에서 매번 recalc를 돌려, 건너뛴 날짜 수만큼 전체 재계산이
 * 반복되어(사실상 O(N^2)) 기록이 쌓일수록 저장이 느려지는 문제가 있었다.
 */
function writeRowRaw_(sheet, dateStr, revenue, accountVal, memo, stockSumAgg, tickerCols, tickerValues) {
  const loc = locateDateRow_(sheet, dateStr);
  let targetRow;
  if (loc.row > 0) {
    targetRow = loc.row;
  } else {
    targetRow = loc.insertBeforeRow;
    // 🩹 [정밀검사 개선] 이 날짜가 기존 기록들보다 과거인데 그 자리에 행이 없는 경우
    // (진짜 구멍난 날짜), 시트 맨 끝이 아니라 정확한 날짜순 위치에 빈 행을 끼워 넣는다.
    // insertBeforeRow가 "현재 마지막 행 + 1"(맨 끝에 추가하는 경우)이면 끼워 넣을
    // 필요 없이 그냥 새 행에 쓰면 된다.
    if (targetRow <= sheet.getLastRow()) {
      sheet.insertRowBefore(targetRow);
    }
  }

  const cashCol = ensureCashColumn_(sheet);
  const maxCol = Math.max(sheet.getLastColumn(), cashCol);

  let rowData = new Array(maxCol).fill('');
  // 기존 행이면 현재 값을 기반으로 덮어써서, 관리하지 않는 열(사용자 추가 열/중복 헤더 열)이 지워지지 않게 한다.
  if (loc.row > 0) rowData = sheet.getRange(targetRow, 1, 1, maxCol).getValues()[0];
  rowData[MAIN_COLS.date - 1] = dateStr;
  rowData[MAIN_COLS.revenue - 1] = revenue;
  rowData[MAIN_COLS.accountVal - 1] = accountVal;
  rowData[MAIN_COLS.memo - 1] = memo || '';
  rowData[MAIN_COLS.stockSumAgg - 1] = stockSumAgg;

  Object.keys(tickerCols).forEach(t => {
    const col = tickerCols[t];
    rowData[col - 1] = tickerValues[t] !== undefined ? tickerValues[t] : 0;
  });

  const cash = accountVal - stockSumAgg;
  rowData[cashCol - 1] = cash;

  // 🩹 [정밀검사 개선] 그리드 시트들(매수목록/차수기록/메모장 등)과 동일하게, 날짜 열도
  // setValues로 쓰기 전에 텍스트 서식('@')을 먼저 적용한다. 이걸 안 하면 "2026-01-15" 같은
  // 문자열이 구글시트에 의해 Date 직렬값으로 자동 변환될 수 있고, 그러면 locateDateRow_의
  // 이진 탐색(문자열 오름차순 비교 전제)이 깨지거나 formatDate_ 재변환 과정에서 스크립트
  // 타임존과 실제 사용자 타임존이 다를 경우 날짜가 하루 밀릴 수 있다.
  sheet.getRange(targetRow, MAIN_COLS.date, 1, 1).setNumberFormat('@');

  // 기본 필드 일괄 쓰기 (파생 컬럼은 비워두고 호출부에서 나중에 다시 계산)
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  return targetRow;
}

/**
 * 🩹 [정밀검사 개선] saveDailyRecord_의 자동채움 구간 전용 배치 쓰기 함수.
 * gapDates는 전부 시트의 현재 마지막 행(dateStr)보다 미래이고 아직 시트에 존재하지 않는
 * 날짜들이라는 전제 하에, locateDateRow_(전체 날짜 열 재탐색)를 날짜 수만큼 반복하지 않고
 * "현재 마지막 행 바로 다음 행"부터 순서대로 이어 붙일 위치를 한 번만 계산한다.
 * 모든 공백일이 같은 계좌평가액/종목합산평가금액/종목별 값(직전 마지막 기록값 그대로)을
 * 공유하므로, 행 데이터를 메모리에서 전부 만든 뒤 setNumberFormat + setValues를 각각
 * 딱 한 번씩만 호출해 시트 API 왕복 횟수를 날짜 수와 무관하게 일정하게 유지한다.
 * 반환값은 배치로 쓴 행들 중 가장 앞쪽(가장 오래된 공백일) 행 번호로, 호출부가
 * earliestChangedRow를 갱신하는 데 그대로 쓰인다.
 */
function writeGapRowsBatch_(sheet, gapDates, accountVal, stockSumAgg, tickerCols, tickerValues) {
  const cashCol = ensureCashColumn_(sheet);
  const maxCol = Math.max(sheet.getLastColumn(), cashCol);
  const startRow = sheet.getLastRow() + 1;
  const cash = accountVal - stockSumAgg;

  const rows = gapDates.map(dateStr => {
    const rowData = new Array(maxCol).fill('');
    rowData[MAIN_COLS.date - 1] = dateStr;
    rowData[MAIN_COLS.revenue - 1] = 0;
    rowData[MAIN_COLS.accountVal - 1] = accountVal;
    rowData[MAIN_COLS.memo - 1] = AUTO_FILL_MEMO_MARKER;
    rowData[MAIN_COLS.stockSumAgg - 1] = stockSumAgg;
    Object.keys(tickerCols).forEach(t => {
      const col = tickerCols[t];
      rowData[col - 1] = tickerValues[t] !== undefined ? tickerValues[t] : 0;
    });
    rowData[cashCol - 1] = cash;
    return rowData;
  });

  // 🩹 [정밀검사 개선] 공백 날짜가 수백~수천 개(몇 년치)에 달하면 한 번의 setValues에
  // 실려가는 데이터가 지나치게 커져 recalcDerivedFrom_와 동일한 실행시간/응답 크기
  // 제한 위험이 있었다. recalcDerivedFrom_와 동일하게 RECALC_CHUNK_SIZE_행씩 나눠서
  // setNumberFormat/setValues를 호출해, 한 번에 오가는 데이터 양을 일정하게 유지한다
  // (결과는 청크로 나누기 전과 완전히 동일 - 각 행의 값은 서로 독립적이라 청크 경계에서
  // 이어받을 상태가 없음).
  for (let chunkStart = 0; chunkStart < rows.length; chunkStart += RECALC_CHUNK_SIZE_) {
    const chunkRows = rows.slice(chunkStart, chunkStart + RECALC_CHUNK_SIZE_);
    const sheetRowStart = startRow + chunkStart;
    sheet.getRange(sheetRowStart, MAIN_COLS.date, chunkRows.length, 1).setNumberFormat('@');
    sheet.getRange(sheetRowStart, 1, chunkRows.length, maxCol).setValues(chunkRows);
  }

  return startRow;
}

/**
 * ⚠️ [정밀검사 메모] 여기서 쓰는 누적수익금/MDD/일별수익률 계산식은 index.html의
 * parseAndRenderData()가 프런트에서 다시 한 번 독립적으로 계산하는 식과 반드시 동일해야 한다.
 * 프런트가 이 시트 칸(C/E/F열)을 그대로 신뢰하지 않고 매번 재계산하는 이유는, 시트를 수동으로
 * 편집해 이 값들이 깨졌을 때도 대시보드가 스스로 바로잡아 보여주기 위한 안전장치다. 즉 이 중복은
 * 의도된 것이며, 여기서 계산식을 바꾸면 index.html 쪽도 반드시 같이 바꿔야 두 화면의 숫자가
 * 어긋나지 않는다.
 * startRow부터 마지막 행까지 누적수익금(cumRevenue) / MDD / 일별수익률(dailyReturn)을
 * 순서대로 다시 계산해서 저장한다. startRow보다 앞의 행들은 그대로 두고,
 * startRow 직전 행의 값을 기준(baseline)으로 이어서 계산한다.
 */
// 🩹 [정밀검사 개선 → 구조적 수정] recalcDerivedFrom_는 startRow부터 마지막 행까지
// 다시 읽고 다시 써야 한다(누적수익금/MDD/일별수익률이 전부 "직전 행" 기준 누적 계산이라
// 건너뛸 방법이 없음). 예전에는 이걸 getRange/getValues/setValues 각각 "통째로 한 번"
// 호출로 처리했는데, 데이터가 몇 년 치(수천~수만 행) 쌓인 뒤 아주 오래된 기록을 고치면
// 한 번에 오가는 값 배열이 지나치게 커져 Apps Script 실행시간 제한(6분)에 걸리거나
// 처리 자체가 느려질 위험이 있었다. 이제 RECALC_CHUNK_SIZE_행씩 나눠서 읽고 쓰도록
// 바꿔, 한 번에 다루는 데이터 양을 일정하게 유지한다(전체 계산 결과는 청크로 나누기 전과
// 완전히 동일 - 누적값들을 청크 경계에서도 그대로 이어받아 계산한다). 이렇게 하면:
//  1) 한 번의 getRange 호출이 다루는 셀 수가 항상 일정 수준 이하로 유지되어 처리 시간이
//     안정적이다.
//  2) 만에 하나 실행시간 제한에 걸리더라도, 이미 처리된 앞쪽 청크는 시트에 반영된
//     상태로 남는다(전부 다시 계산해야 하는 대신, 다음 저장/삭제 때 다시 recalc가 돌면서
//     자연히 이어서 바로잡힌다 - 이는 예전의 "한 번에 전부 쓰기" 방식이 보장하던 "부분
//     반영 없음"보다는 약하지만, 그 대신 애초에 시간 제한에 걸릴 위험 자체를 크게 줄인다).
// 이 청크 처리 덕분에 "몇 초면 끝나는 정상 범위"가 훨씬 넓어졌으므로, 아래 경고 기준선도
// 실제 체감 성능에 맞게 상향한다: 배치로 처리하면 수천 행 재계산은 보통 몇 초 안에 끝나므로,
// 2,000행은 지나치게 보수적인 기준이었다. 개인용 일별 기록 앱에서 "정말 오래 걸릴 수 있는"
// 수준인 5,000행으로 올려, 평소에는 불필요한 경고가 뜨지 않게 한다.
const RECALC_WARNING_ROW_THRESHOLD = 5000;
const RECALC_CHUNK_SIZE_ = 2000;

function recalcDerivedFrom_(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  if (startRow > lastRow) return { rowsRecalculated: 0 };

  let maxSoFar = 0;
  if (startRow > 2) {
    const history = sheet.getRange(2, MAIN_COLS.accountVal, startRow - 2, 1).getValues().flat().map(Number);
    history.forEach(v => { if (!isNaN(v) && v > maxSoFar) maxSoFar = v; });
  }

  let prevCum = 0;
  let prevAcc = null;
  if (startRow > 2) {
    prevCum = Number(sheet.getRange(startRow - 1, MAIN_COLS.cumRevenue).getValue()) || 0;
    prevAcc = Number(sheet.getRange(startRow - 1, MAIN_COLS.accountVal).getValue()) || null;
  }

  const totalRows = lastRow - startRow + 1;
  const lastColNeeded = Math.max(sheet.getLastColumn(), MAIN_COLS.dailyReturn);

  for (let chunkStart = startRow; chunkStart <= lastRow; chunkStart += RECALC_CHUNK_SIZE_) {
    const chunkRows = Math.min(RECALC_CHUNK_SIZE_, lastRow - chunkStart + 1);
    const range = sheet.getRange(chunkStart, 1, chunkRows, lastColNeeded);
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      const revenue = Number(values[i][MAIN_COLS.revenue - 1]) || 0;
      const accountVal = Number(values[i][MAIN_COLS.accountVal - 1]) || 0;

      prevCum += revenue;
      if (accountVal > maxSoFar) maxSoFar = accountVal;
      const mdd = maxSoFar > 0 ? (accountVal - maxSoFar) / maxSoFar : 0;
      const dailyReturn = (prevAcc && prevAcc > 0) ? (accountVal - prevAcc) / prevAcc : 0;

      values[i][MAIN_COLS.cumRevenue - 1] = prevCum;
      values[i][MAIN_COLS.mdd - 1] = mdd;
      values[i][MAIN_COLS.dailyReturn - 1] = dailyReturn;

      prevAcc = accountVal;
    }

    range.setValues(values);
  }

  return { rowsRecalculated: totalRows };
}

/**
 * 🆕 deleteDailyRecord_
 * 메인 시트(시트1)에서 특정 날짜의 행을 완전히 삭제한다.
 * 삭제된 행 이후의 행들은 시트에서 한 칸씩 위로 당겨지므로, 그 위치부터
 * 다시 recalcDerivedFrom_로 누적수익금/MDD/일별수익률을 재계산해야
 * (예: 삭제된 날짜 다음 날의 일별수익률이 삭제된 날짜 기준으로 계산된
 * 예전 값 그대로 남는 문제를 방지) 파생 컬럼이 어긋나지 않는다.
 */
function deleteDailyRecord_(sheet, dateStr) {
  dateStr = String(dateStr || '').trim();
  if (!dateStr) throw new Error('date가 누락되었습니다.');

  const row = findDateRow_(sheet, dateStr);
  if (row < 0) throw new Error('해당 날짜의 기록을 찾을 수 없습니다: ' + dateStr);

  snapshotBeforeDestructiveAction_();
  sheet.deleteRow(row);

  // 삭제된 행 자리(이제는 다음 날짜 행이 올라와 있음)부터 끝까지 파생값 재계산.
  // 삭제한 행이 마지막 행이었다면 row가 이제 lastRow보다 커지므로
  // recalcDerivedFrom_ 내부에서 아무 작업 없이 바로 반환된다.
  const recalcResult = recalcDerivedFrom_(sheet, row);
  if (recalcResult && recalcResult.rowsRecalculated > RECALC_WARNING_ROW_THRESHOLD) {
    return { recalcWarning: `과거 기록 삭제로 ${recalcResult.rowsRecalculated.toLocaleString()}행을 다시 계산했습니다. 데이터가 계속 쌓이면 이 작업이 느려지거나 실패할 수 있습니다.` };
  }
  return {};
}

// ===================================================================
// 🗂️ 범용 "그리드" 시트
// ===================================================================

function getOrCreateGridSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 1).setValue('구분');
  }
  return sheet;
}

const MEMO_HEADER = '메모';

function normalizeCellText_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '').trim();
}

function normalizeCellValue_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (val === null || val === undefined) return '';
  return val;
}

// 🩹 [정밀검사 개선] 중복 종목 헤더 방어 (그리드 시트 버전). 메인 시트의
// getTickerColumns_와 동일한 문제: 시트를 직접 편집해 같은 종목명 헤더가 두 열
// 이상 생기면, 예전에는 result.tickers 배열에 같은 이름이 중복으로 들어가고
// data[rowLabel][티커] 값은 나중 열이 앞선 열을 조용히 덮어썼다(값 손실은 아니고
// "안 보이게" 되는 정도이지만, 화면에는 같은 종목이 두 번 나타날 수 있었다).
// 이제 메인 시트와 동일한 규칙(나중 열 우선)으로 명시적으로 합치고, tickers
// 목록에는 종목명당 한 번만 넣으며, 중복이 발견되면 실행 로그에 경고를 남긴다.
function readGrid_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const result = { rowLabels: [], tickers: [], data: {}, memos: {} };

  const colToTicker = {};
  const tickerSeenCol = {};
  let memoCol = -1;

  if (lastCol >= 2) {
    const headers = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    headers.forEach((h, idx) => {
      const col = idx + 2;
      const t = normalizeCellText_(h);
      if (!t) return;
      if (t === MEMO_HEADER) {
        memoCol = col;
      } else {
        colToTicker[col] = t;
        if (tickerSeenCol[t] !== undefined) {
          Logger.log(`⚠️ 그리드 시트(${sheet.getName()})에 종목 헤더 "${t}"가 ${tickerSeenCol[t]}열과 ${col}열에 중복 존재합니다. ` +
            `${tickerSeenCol[t]}열의 값은 조회에서 무시되고 ${col}열만 사용됩니다. cleanupDuplicateTickerHeaders_()로 정리할 수 있습니다.`);
        } else {
          result.tickers.push(t);
        }
        tickerSeenCol[t] = col;
      }
    });
  }

  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    rows.forEach(r => {
      const rowLabel = normalizeCellText_(r[0]);
      if (!rowLabel) return;
      result.rowLabels.push(rowLabel);
      result.data[rowLabel] = {};
      result.memos[rowLabel] = '';
      for (let i = 1; i < r.length; i++) {
        const col = i + 1;
        if (col === memoCol) {
          result.memos[rowLabel] = normalizeCellValue_(r[i]);
        } else if (colToTicker[col]) {
          result.data[rowLabel][colToTicker[col]] = normalizeCellValue_(r[i]);
        }
      }
    });
  }

  return result;
}

function ensureMemoColumn_(sheet) {
  return addGridTickerColumn_(sheet, MEMO_HEADER);
}

// 🩹 [정밀검사 개선] 예전에는 첫 번째로 일치하는 열을 반환했는데, readGrid_/
// getTickerColumns_는 "나중 열 우선"으로 값을 읽으므로 중복 헤더가 있을 때
// addGridTickerColumn_/deleteGridTickerColumn_ 등 이 함수를 쓰는 쓰기 경로가 실제
// 데이터가 읽히는 열과 다른 열(더 앞선, 이미 무시되고 있는 열)을 조작하는 불일치가
// 있었다. 마지막으로 일치하는 열을 반환하도록 통일한다.
function findGridTickerCol_(sheet, ticker) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return -1;
  const headers = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  const wanted = String(ticker).trim().toUpperCase();
  let found = -1;
  for (let i = 0; i < headers.length; i++) {
    if (normalizeCellText_(headers[i]).toUpperCase() === wanted) {
      found = i + 2;
    }
  }
  return found;
}

function addGridTickerColumn_(sheet, ticker) {
  ticker = String(ticker).trim();
  if (!ticker) return -1;
  const existing = findGridTickerCol_(sheet, ticker);
  if (existing > 0) return existing;
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setNumberFormat('@').setValue(ticker);
  return newCol;
}

function deleteGridTickerColumn_(sheet, ticker) {
  const col = findGridTickerCol_(sheet, ticker);
  if (col > 0) sheet.deleteColumn(col);
}

function findGridRow_(sheet, rowLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const labels = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < labels.length; i++) {
    if (normalizeCellText_(labels[i][0]) === rowLabel) return i + 2;
  }
  return -1;
}

function addGridRow_(sheet, rowLabel) {
  rowLabel = String(rowLabel).trim();
  if (!rowLabel) return -1;
  const existing = findGridRow_(sheet, rowLabel);
  if (existing > 0) return existing;
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setNumberFormat('@').setValue(rowLabel);
  return newRow;
}

function deleteGridRow_(sheet, rowLabel) {
  const row = findGridRow_(sheet, rowLabel);
  if (row > 0) sheet.deleteRow(row);
}

function saveGridCell_(sheet, rowLabel, ticker, value) {
  const col = addGridTickerColumn_(sheet, ticker);
  const row = addGridRow_(sheet, rowLabel);
  const cell = sheet.getRange(row, col);
  cell.setNumberFormat('@');
  cell.setValue(value === undefined ? '' : value);
}

/**
 * 🆕 saveGridRow_
 * 한 행(rowLabel)에 대해 여러 칸(valuesMap = { 컬럼명: 값, ... })을
 * 한 번의 시트 접근으로 저장한다. 프런트에서 필드별로 saveGridCell을
 * 동시에 여러 번 호출하던 것을 이 액션 하나로 대체해, 락 경합으로 인한
 * 무작위 저장 실패를 줄인다.
 */
function saveGridRow_(sheet, rowLabel, valuesMap) {
  const row = addGridRow_(sheet, rowLabel);
  Object.keys(valuesMap).forEach(ticker => {
    const col = addGridTickerColumn_(sheet, ticker);
    const cell = sheet.getRange(row, col);
    cell.setNumberFormat('@');
    const value = valuesMap[ticker];
    cell.setValue(value === undefined ? '' : value);
  });
}

// ===================================================================
// 📌 종목별 매수 관리 "확정 내역" (기록 추가 버튼으로 영구 저장되는 수익금 기록)
// ===================================================================
// '종목별 매수 관리'의 할당($)/매수($) 칸은 언제든 수정될 수 있는 "작업용" 값이다.
// 사용자가 특정 계좌·종목의 실현 수익금을 확정하고 싶을 때, 그 칸의 💰 토글을 열어
// 날짜/메모/수익금을 입력하고 '기록 추가'를 누르면 그 시점의 값을 이 시트에
// 별도의 행으로 영구 저장한다. 이렇게 저장된 값은 "당일 수익금" 자동 합산(반자동 채우기)의
// 근거 데이터가 되며, 원본 할당/매수 칸을 나중에 지우거나 바꿔도 그대로 남는다.
//
// (참고: 예전에는 '차수 기록'에도 같은 방식의 수익금 확정 기능이 있었으나,
//  같은 매매 건이 두 표에 중복 확정되어 "당일 수익금" 합산이 두 배로 잡히는 문제를 막기 위해
//  수익 확정은 이 '종목별 매수 관리' 한 곳으로 일원화했다. '차수 기록'은 순수 기록용으로만 남는다.)

const BUYMGMT_RECORDS_SHEET = '매수관리_확정';
const BUYMGMT_RECORDS_HEADERS = ['ID', '날짜', '종목', '계좌', '메모', '수익금', '등록시각'];

function getOrCreateBuyMgmtRecordsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BUYMGMT_RECORDS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BUYMGMT_RECORDS_SHEET);
    sheet.getRange(1, 1, 1, BUYMGMT_RECORDS_HEADERS.length).setValues([BUYMGMT_RECORDS_HEADERS]);
    sheet.setColumnWidths(1, BUYMGMT_RECORDS_HEADERS.length, 110);
  }
  return sheet;
}

function readBuyMgmtRecords_(sheet) {
  const lastRow = sheet.getLastRow();
  const records = [];
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, BUYMGMT_RECORDS_HEADERS.length).getValues();
    rows.forEach(r => {
      const id = String(r[0] || '').trim();
      if (!id) return;
      records.push({
        id: id,
        date: normalizeCellValue_(r[1]),
        ticker: String(r[2] || ''),
        account: String(r[3] || ''),
        memo: String(r[4] || ''),
        profit: (typeof r[5] === 'number') ? r[5] : (Number(String(r[5] === null || r[5] === undefined ? '' : r[5]).replace(/[$,\s]/g, '')) || 0),
        createdAt: normalizeCellValue_(r[6])
      });
    });
  }
  return { records: records };
}

function findBuyMgmtRecordRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return -1;
}

function addBuyMgmtRecord_(sheet, rec) {
  const id = 'r' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000);
  // 🩹 [정밀검사 개선] 메인 시트(시트1)의 날짜 열과 동일한 문제: 날짜 문자열을
  // appendRow로 그대로 쓰면 구글시트가 Date 타입으로 자동 변환할 수 있고, 스크립트
  // 타임존과 사용자 타임존이 다르면 읽어올 때 하루가 밀릴 수 있다. B열(날짜)에
  // 텍스트 서식을 먼저 적용해 막는다.
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 2, 1, 1).setNumberFormat('@');
  sheet.getRange(newRow, 1, 1, 7).setValues([[
    id,
    rec.date || '',
    String(rec.ticker || '').trim(),
    rec.account || '',
    rec.memo || '',
    rec.profit === undefined || rec.profit === '' ? '' : Number(rec.profit),
    new Date()
  ]]);
  return id;
}

function updateBuyMgmtRecord_(sheet, id, rec) {
  const row = findBuyMgmtRecordRow_(sheet, id);
  if (row < 0) throw new Error('수정할 기록을 찾을 수 없습니다: ' + id);
  // 🩹 [정밀검사 개선] deleteBuyMgmtRecord_는 삭제 직전 스냅샷을 남기지만, 이 함수(수정)는
  // 기존 값을 그 자리에서 덮어써 이전 값을 되돌릴 방법이 없었다(삭제와 마찬가지로 파괴적
  // 작업). 잘못 수정해도 그날의 자동 백업(하루 1회)으로 복구할 수 있도록 동일하게 스냅샷한다.
  snapshotBeforeDestructiveAction_();
  sheet.getRange(row, 2, 1, 1).setNumberFormat('@');
  sheet.getRange(row, 2, 1, 5).setValues([[
    rec.date || '',
    String(rec.ticker || '').trim(),
    rec.account || '',
    rec.memo || '',
    rec.profit === undefined || rec.profit === '' ? '' : Number(rec.profit)
  ]]);
}

function deleteBuyMgmtRecord_(sheet, id) {
  const row = findBuyMgmtRecordRow_(sheet, id);
  if (row > 0) {
    snapshotBeforeDestructiveAction_();
    sheet.deleteRow(row);
  }
}

// ===================================================================
// 🧹 중복 행 정리 도구
// ===================================================================

function dedupeGridRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const merged = {};
  const order = [];

  rows.forEach(r => {
    const label = normalizeCellText_(r[0]);
    if (!label) return;
    if (!merged[label]) {
      merged[label] = r.slice();
      order.push(label);
    } else {
      for (let i = 1; i < r.length; i++) {
        if (r[i] !== '' && r[i] !== null && r[i] !== undefined) {
          merged[label][i] = r[i];
        }
      }
    }
    merged[label][0] = label;
  });

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  const newRows = order.map(label => merged[label]);
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, lastCol).setValues(newRows);
    sheet.getRange(2, 1, newRows.length, 1).setNumberFormat('@');
  }
}

function cleanupBuyListDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('매수목록');
  if (sheet) dedupeGridRows_(sheet);
}

function cleanupStepLogDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('차수기록');
  if (sheet) dedupeGridRows_(sheet);
}

function cleanupNotepadDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('메모장');
  if (sheet) dedupeGridRows_(sheet);
}

// 🩹 [정밀검사 개선] 실제로 쓰이는 메모장1/수익률데이터에는 정리 함수가 없어서
// 중복 행이 생겨도 되돌릴 방법이 없었다. 다른 그리드들과 동일한 패턴으로 추가.
function cleanupNotepad1Duplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('메모장1');
  if (sheet) dedupeGridRows_(sheet);
}

function cleanupYieldDataDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('수익률데이터');
  if (sheet) dedupeGridRows_(sheet);
}

// 🩹 [정밀검사 개선] cleanupTradePlanDuplicates()('매매예정' 시트 중복 정리 유틸)는
// 그 대상 시트인 '매매예정' 자체가 프런트엔드 어디에서도 참조되지 않는 완전한 죽은
// 기능이었다(과거 한때 중복 선언까지 있었던 코드). 실제로 쓰이지 않는 시트의 정리
// 유틸리티를 유지할 이유가 없어 함수 자체를 제거했다. 다른 그리드(매수목록/차수기록/
// 메모장/메모장1/수익률데이터)의 중복 정리 유틸은 실제로 쓰는 시트들이므로 그대로 둔다.

// ===================================================================
// 🩹 [정밀검사 개선] 메인 시트(시트1) 중복 날짜 정리 도구
// ===================================================================
// 정상적인 저장 흐름(saveDailyRecord_ → locateDateRow_)은 upsert 방식이라 같은 날짜가
// 두 번 저장될 수 없지만, 사용자가 구글시트를 직접 열어 행을 복사/붙여넣기하면 같은
// 날짜가 중복으로 존재할 수 있다. readMainData_/parseAndRenderData는 방어적으로 응답 단계에서
// 중복을 걸러내지만, 시트 자체의 중복 행은 그대로 남아있으므로 근본적으로는 이 함수로
// 한 번 정리해줘야 한다. 다른 그리드 시트들의 dedupeGridRows_와 동일한 규칙(먼저 나온 행을
// 기준으로, 나중 중복 행의 비어있지 않은 값으로 덮어써 병합)으로 중복을 제거한 뒤, 날짜
// 오름차순으로 재정렬해서 다시 쓰고, 파생 컬럼(누적수익금/MDD/일별수익률)을 처음부터 다시
// 계산한다.
//
// [사용 방법] 이 편집기 상단 함수 목록에서 "cleanupMainSheetDuplicateDates_" 선택 → ▶ 실행.
// 시트 내용을 되돌릴 수 없게 바꾸는 작업이므로, 실행 전 runDailyBackup_()으로 백업을
// 한 번 남겨두는 것을 권장한다(이 함수 자체도 실행 전 자동으로 스냅샷을 한 번 남긴다).
function cleanupMainSheetDuplicateDates_() {
  const sheet = getMainSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('정리할 데이터가 2행 이하입니다. 중복이 있을 수 없습니다.');
    return { duplicatesRemoved: 0 };
  }

  const lastCol = Math.max(sheet.getLastColumn(), MAIN_COLS.tickerStart - 1);
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const merged = {};
  const order = [];
  let duplicatesRemoved = 0;

  rows.forEach(r => {
    const dateStr = formatDate_(r[MAIN_COLS.date - 1]);
    if (!dateStr) return;
    if (!merged[dateStr]) {
      merged[dateStr] = r.slice();
      order.push(dateStr);
    } else {
      duplicatesRemoved++;
      // 나중 중복 행의 비어있지 않은 값으로 덮어쓴다(다른 그리드 정리 유틸과 동일한 병합 규칙 -
      // 즉 각 열마다 "가장 최근에 값이 채워진" 쪽이 최종값으로 남는다).
      for (let i = 0; i < r.length; i++) {
        if (r[i] !== '' && r[i] !== null && r[i] !== undefined) {
          merged[dateStr][i] = r[i];
        }
      }
      merged[dateStr][MAIN_COLS.date - 1] = dateStr;
    }
  });

  if (duplicatesRemoved === 0) {
    Logger.log('중복된 날짜가 없습니다. 정리할 내용이 없습니다.');
    return { duplicatesRemoved: 0 };
  }

  snapshotBeforeDestructiveAction_();

  // 날짜 오름차순으로 재정렬해, locateDateRow_의 이진탐색이 전제하는 정렬 상태를 다시 보장한다.
  order.sort();
  const newRows = order.map(dateStr => merged[dateStr]);

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  sheet.getRange(2, 1, newRows.length, lastCol).setValues(newRows);
  sheet.getRange(2, MAIN_COLS.date, newRows.length, 1).setNumberFormat('@');

  const recalcResult = recalcDerivedFrom_(sheet, 2);
  invalidateBootstrapCache_();

  Logger.log(`중복 날짜 ${duplicatesRemoved}건을 정리했습니다. (병합 후 총 ${recalcResult.rowsRecalculated}행 재계산)`);
  return { duplicatesRemoved: duplicatesRemoved, rowsRecalculated: recalcResult.rowsRecalculated };
}

// ===================================================================
// 🩹 [정밀검사 개선] 중복 종목 헤더 정리 도구 (메인 시트 + 그리드 시트 공통)
// ===================================================================
// 시트를 직접 편집하다 실수로 같은 종목명 헤더가 두 열 이상 생기면, getTickerColumns_/
// readGrid_는 시트에서 더 뒤에 있는 열의 값만 읽고 앞선 열의 값은 조회에서 조용히
// 무시한다(값 자체가 지워지진 않지만 화면에는 보이지 않게 된다). 지난번 "중복 날짜"
// 정리 도구(cleanupMainSheetDuplicateDates_)와 동일한 방식으로, 중복된 종목 열들을
// 하나로 합친다: 각 데이터 행마다 더 뒤에 있는(=더 최근에 편집됐을 가능성이 높은)
// 열의 비어있지 않은 값을 우선하고, 그 열이 비어있으면 앞선 열의 값을 그대로 남겨
// 데이터가 사라지지 않게 병합한 뒤, 남는 중복 열은 삭제한다.
//
// [사용 방법] 이 편집기 상단 함수 목록에서 "cleanupDuplicateTickerHeaders_" 선택 →
// ▶ 실행. 메인 시트(시트1)와 그리드 시트(매수목록/차수기록/메모장/메모장1/수익률데이터)
// 전체를 한 번에 검사해 정리한다. 열을 삭제하는 되돌릴 수 없는 작업이므로, 실행 전
// runDailyBackup_()으로 백업을 한 번 남겨두는 것을 권장한다(이 함수 자체도 실행 전
// 자동으로 스냅샷을 한 번 남긴다).
function mergeDuplicateHeaderColumns_(sheet, headerRow, headerStartCol, dataStartRow, extraSkipHeaders) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < headerStartCol) return { duplicatesRemoved: 0 };

  const skipHeaders = (extraSkipHeaders || []).map(h => String(h).trim().toUpperCase());
  const numHeaderCols = lastCol - headerStartCol + 1;
  const headers = sheet.getRange(headerRow, headerStartCol, 1, numHeaderCols).getValues()[0];

  // 종목명(대문자로 정규화) -> 그 이름을 가진 모든 열 번호(오름차순, 시트에 나온 순서)
  const nameToCols = {};
  headers.forEach((h, idx) => {
    const t = String(h || '').trim().toUpperCase();
    if (!t || skipHeaders.indexOf(t) >= 0) return;
    const col = headerStartCol + idx;
    if (!nameToCols[t]) nameToCols[t] = [];
    nameToCols[t].push(col);
  });

  const dupGroups = Object.keys(nameToCols).map(t => nameToCols[t]).filter(cols => cols.length > 1);
  if (dupGroups.length === 0) return { duplicatesRemoved: 0 };

  const numDataRows = Math.max(0, lastRow - dataStartRow + 1);
  let duplicatesRemoved = 0;
  const colsToDelete = [];

  dupGroups.forEach(cols => {
    const keepCol = cols[0]; // 가장 앞선 열의 "위치"를 남기고, 값은 더 뒤 열을 우선해 병합
    if (numDataRows > 0) {
      const keepValues = sheet.getRange(dataStartRow, keepCol, numDataRows, 1).getValues();
      for (let c = 1; c < cols.length; c++) {
        const otherValues = sheet.getRange(dataStartRow, cols[c], numDataRows, 1).getValues();
        for (let r = 0; r < numDataRows; r++) {
          const v = otherValues[r][0];
          if (v !== '' && v !== null && v !== undefined) keepValues[r][0] = v;
        }
      }
      sheet.getRange(dataStartRow, keepCol, numDataRows, 1).setValues(keepValues);
    }
    duplicatesRemoved += cols.length - 1;
    colsToDelete.push(...cols.slice(1));
  });

  // 열 삭제는 뒤(큰 열 번호)에서부터 해야 앞선 열 번호가 밀려서 꼬이지 않는다.
  colsToDelete.sort((a, b) => b - a).forEach(col => sheet.deleteColumn(col));

  return { duplicatesRemoved: duplicatesRemoved };
}

function cleanupDuplicateTickerHeaders_() {
  snapshotBeforeDestructiveAction_();

  let totalRemoved = 0;
  const details = [];

  const mainSheet = getMainSheet_();
  const mainResult = mergeDuplicateHeaderColumns_(mainSheet, 1, MAIN_COLS.tickerStart, 2, [CASH_HEADER]);
  if (mainResult.duplicatesRemoved > 0) {
    details.push(`${mainSheet.getName()}: ${mainResult.duplicatesRemoved}개 열`);
    totalRemoved += mainResult.duplicatesRemoved;
  }

  Object.values(BOOTSTRAP_GRID_SHEETS).forEach(sheetName => {
    const gridSheet = getOrCreateGridSheet_(sheetName);
    const result = mergeDuplicateHeaderColumns_(gridSheet, 1, 2, 2, [MEMO_HEADER]);
    if (result.duplicatesRemoved > 0) {
      details.push(`${sheetName}: ${result.duplicatesRemoved}개 열`);
      totalRemoved += result.duplicatesRemoved;
    }
  });

  if (totalRemoved === 0) {
    Logger.log('중복된 종목 헤더가 없습니다. 정리할 내용이 없습니다.');
    return { duplicatesRemoved: 0 };
  }

  invalidateBootstrapCache_();
  Logger.log(`중복 종목 헤더 총 ${totalRemoved}개 열을 정리했습니다. (${details.join(', ')})`);
  return { duplicatesRemoved: totalRemoved, details: details };
}

// ===================================================================
// 🩹 [정밀검사 개선] 메인 시트 날짜 열(A열) 서식 복구용 1회 실행 유틸리티
// ===================================================================
// writeRowRaw_에 텍스트 서식 적용이 빠져 있던 동안 저장된 기존 행들 중 일부가
// 구글시트에 의해 Date 타입으로 자동 변환되어 있을 수 있다. 이 함수를 스크립트
// 편집기에서 한 번 수동 실행하면: A열 전체를 읽어 formatDate_로 문자열로 정규화한
// 뒤 다시 쓰고, 열 전체에 텍스트 서식('@')을 적용해 앞으로도 자동 변환이 일어나지
// 않게 막는다. (listBackupFiles_ 등 기존 "수동 1회 실행" 유틸리티와 동일한 패턴.)
function fixMainDateColumnFormat_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시트1');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, MAIN_COLS.date, lastRow - 1, 1);
  const values = range.getValues();
  const normalized = values.map(r => [formatDate_(r[0])]);

  range.setNumberFormat('@');
  range.setValues(normalized);
}

// ===================================================================
// 🩹 [정밀검사 개선] 종목 헤더/티커 셀 텍스트 서식 복구용 1회 실행 유틸리티
// ===================================================================
// addTicker_/saveStep_/ensureTickerColumn_에 텍스트 서식 적용이 빠져 있던 동안
// "005930" 같은 순수 숫자 티커(한국 종목코드 등)를 추가했다면, 구글시트가 이를
// 숫자(5930)로 자동 변환해 앞자리 0이 사라졌을 수 있다. 이 함수를 한 번 실행하면:
// 분할매수데이터(A열/F열)와 메인 시트(시트1, I열~ 헤더)의 값을 다시 문자열로
// 정규화해 쓰고 텍스트 서식('@')을 적용한다.
// ⚠️ 주의: 이미 "5930"처럼 숫자로 굳어진 값은 원래 앞자리 0을 복원할 방법이 없다
// (예: 5930 → "5930", 원래 "005930"이었는지는 알 수 없음). 이 함수는 "앞으로 더 이상
// 숫자로 변환되지 않게" 막아줄 뿐이므로, 이미 깨진 티커가 있다면 시트에서 직접
// 올바른 값으로 한 번 고쳐 쓴 뒤 실행하는 것을 권장한다.
function fixTickerHeaderFormats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const splitBuySheet = ss.getSheetByName(SHEET_NAME);
  if (splitBuySheet) {
    const lastRow = splitBuySheet.getLastRow();
    if (lastRow >= 2) {
      const colA = splitBuySheet.getRange(2, 1, lastRow - 1, 1);
      colA.setNumberFormat('@');
      colA.setValues(colA.getValues().map(r => [String(r[0] || '').trim().toUpperCase()]));

      const colF = splitBuySheet.getRange(2, 6, lastRow - 1, 1);
      colF.setNumberFormat('@');
      colF.setValues(colF.getValues().map(r => [String(r[0] || '').trim().toUpperCase()]));
    }
  }

  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (mainSheet) {
    const lastCol = mainSheet.getLastColumn();
    if (lastCol >= MAIN_COLS.tickerStart) {
      const headerRange = mainSheet.getRange(1, MAIN_COLS.tickerStart, 1, lastCol - MAIN_COLS.tickerStart + 1);
      headerRange.setNumberFormat('@');
      headerRange.setValues([headerRange.getValues()[0].map(h => String(h || '').trim().toUpperCase())]);
    }
  }

  Object.values(BOOTSTRAP_GRID_SHEETS).forEach(sheetName => {
    const gridSheet = ss.getSheetByName(sheetName);
    if (!gridSheet) return;
    const lastCol = gridSheet.getLastColumn();
    if (lastCol >= 2) {
      const headerRange = gridSheet.getRange(1, 2, 1, lastCol - 1);
      headerRange.setNumberFormat('@');
      headerRange.setValues([headerRange.getValues()[0].map(h => String(h || '').trim())]);
    }
  });
}

// ===================================================================
// 💾 [기능 추가] 자동 백업 (매일 1회, 구글 드라이브에 스냅샷 저장)
// ===================================================================
// 실수로 행/열을 대량 삭제하거나 잘못 편집했을 때 되돌릴 방법이 없었다.
// 스프레드시트의 모든 시트를 매일 한 번 JSON 스냅샷으로 구글 드라이브에
// 저장해두면, 문제가 생겼을 때 그 파일을 열어 값을 다시 채워 넣을 수 있다.
//
// [설치 방법]
// 1) 이 Apps Script 편집기 상단에서 함수 목록을 "createDailyBackupTrigger"로 선택
// 2) ▶ 실행 버튼 클릭 → 최초 1회 권한 승인(구글 드라이브 접근 허용)
// 3) 이후 매일 새벽 4시(스크립트 타임존 기준)에 자동으로 백업이 생성된다.
// 4) 트리거를 끄고 싶으면 함수 목록에서 "deleteDailyBackupTrigger"를 실행.
//
// 백업 파일은 내 드라이브에 "투자대시보드_백업" 폴더가 자동 생성되어 그 안에 쌓이며,
// 오래된 파일은 최근 BACKUP_KEEP_COUNT(60)개만 남기고 자동으로 휴지통으로 이동한다(용량 관리).
// ===================================================================

const BACKUP_FOLDER_NAME = '투자대시보드_백업';
const BACKUP_TRIGGER_HANDLER = 'runDailyBackup_';
const BACKUP_KEEP_COUNT = 60;

function getOrCreateBackupFolder_() {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * 🆕 [기능 추가 - JSON 백업 내보내기/불러오기] 스프레드시트에 있는 모든 시트를 값
 * 그대로(수식 결과값) JSON 스냅샷 객체로 만든다. 기존에는 runDailyBackup_ 안에서
 * 이 로직을 바로 만들고 곧장 드라이브에 저장했는데, exportSnapshot 읽기 액션에서도
 * (드라이브에 저장하지 않고) 동일한 스냅샷을 그대로 응답으로 내려줄 수 있어야 해서
 * 스냅샷 "생성" 부분만 별도 함수로 분리했다.
 */
function buildSnapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const snapshot = {
    backedUpAt: new Date().toISOString(),
    spreadsheetName: ss.getName(),
    sheets: {},
    // 🩹 [정밀검사 개선] 값만 저장하던 예전 방식은 수식을 복원하지 못했다. 수식이 있는 칸은
    // [행, 열, 수식] 목록으로 함께 기록한다(없으면 키 자체를 생략).
    formulas: {}
  };

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    // 🩹 [정밀검사 개선] 비어 있는 시트도 빈 배열로 기록한다. 예전에는 통째로 빠져서,
    // 복원해도 그 시트의 "현재 내용"이 그대로 남았다.
    if (lastRow < 1 || lastCol < 1) { snapshot.sheets[name] = []; return; }
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    snapshot.sheets[name] = range.getValues().map(row =>
      row.map(v => (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v)
    );
    try {
      const formulas = range.getFormulas();
      const list = [];
      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          if (formulas[r][c]) list.push([r + 1, c + 1, formulas[r][c]]);
        }
      }
      if (list.length > 0) snapshot.formulas[name] = list;
    } catch (e) { /* 수식 수집 실패 시 값만 저장 */ }
  });

  return snapshot;
}

/**
 * 스프레드시트 전체 스냅샷을 만들어 드라이브에 저장한다. 트리거로 자동 실행되거나,
 * 편집기에서 수동으로 실행해도 된다.
 */
function runDailyBackup_() {
  // 🩹 [정밀검사 개선] 예약 트리거로 실행될 때는 락 없이 시트를 읽어, 저장(쓰기) 도중의
  // 어중간한 상태가 스냅샷에 찍힐 수 있었다. doPost가 이미 락을 잡은 경로(SCRIPT_LOCK_HELD_)가
  // 아니면 여기서 직접 락을 잡고 스냅샷을 만든다.
  let ownLock = null;
  if (!SCRIPT_LOCK_HELD_) {
    ownLock = LockService.getScriptLock();
    ownLock.waitLock(30000); // 못 잡으면 예외 → 백업 실패로 그대로 드러난다
  }
  try {
    const folder = getOrCreateBackupFolder_();
    const snapshot = buildSnapshot_();

    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
    const fileName = `백업_${dateStr}.json`;
    folder.createFile(fileName, JSON.stringify(snapshot), MimeType.PLAIN_TEXT);

    // 정리 실패(휴지통 이동 오류 등)가 "백업 실패"로 오인되지 않도록 분리한다.
    try { cleanupOldBackups_(folder, BACKUP_KEEP_COUNT); } catch (cleanErr) { /* 무시 */ }
  } finally {
    if (ownLock) ownLock.releaseLock();
  }
}

// ===================================================================
// 🆕 [기능 추가] 파괴적 작업 전 자동 스냅샷 (하루 1회 제한)
// ===================================================================
// deleteDailyRecord_/deleteTickerColumnFromMain_/deleteBuyMgmtRecord_는 confirm()
// 창 하나만 거치면 바로 되돌릴 수 없는 삭제가 실행됐다. runDailyBackup_()이 이미
// 구현돼 있으니, 이 세 액션 직전에 가벼운 스냅샷을 한 번 찍어두면 실수로 삭제했을 때
// listBackupFiles_/restoreFromBackupFile_로 복구할 수 있다. 단, 하루에 여러 번
// 삭제해도 스냅샷은 하루 1번만 만들어서(스크립트 속성에 날짜를 기록) 백업 파일이
// 무한정 쌓이지 않게 한다(어차피 cleanupOldBackups_가 BACKUP_KEEP_COUNT(60)개까지만 보관하긴 하지만,
// 그 60개가 며칠치인지 예측 가능하게 유지하기 위함).
const DESTRUCTIVE_SNAPSHOT_PROP_KEY_ = 'destructiveSnapshotDate';

function snapshotBeforeDestructiveAction_() {
  try {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(DESTRUCTIVE_SNAPSHOT_PROP_KEY_) === todayStr) return; // 오늘 이미 스냅샷 찍음
    runDailyBackup_();
    props.setProperty(DESTRUCTIVE_SNAPSHOT_PROP_KEY_, todayStr);
  } catch (e) {
    // 드라이브 권한 미승인(자동 백업 트리거를 아직 설정 안 한 경우) 등으로 스냅샷이
    // 실패해도, 사용자가 요청한 삭제 자체는 막지 않는다(호출부에서 계속 진행).
  }
}

// 🩹 [정밀검사 개선 - 문제 1] snapshotBeforeDestructiveAction_는 하루 1회 제한이 있어,
// 같은 날 먼저 다른 삭제 작업(일일 기록/종목/확정 기록 삭제 등)이 스냅샷을 이미
// 찍어둔 상태라면 그 이후 "백업에서 복원"을 실행해도 새 백업을 만들지 않고 조용히
// 건너뛴다. 그러면 대시보드 확인창의 "복원 직전 상태는 자동으로 한 번 더 백업됩니다"
// 안내와 달리, 그날 입력한 데이터가 아무 백업에도 남지 않은 채 복원으로 덮어써져
// 영구히 사라질 수 있다. 복원은 그 시점 상태를 반드시 새로 남겨야 확인창의 약속이
// 지켜지므로, 하루 1회 제한 없이 항상 백업하는 전용 함수를 따로 둔다. 백업 후에는
// DESTRUCTIVE_SNAPSHOT_PROP_KEY_도 함께 갱신해, 같은 날 이어지는 삭제 작업들이
// 불필요하게 중복 백업을 또 만들지 않게 한다.
function snapshotBeforeRestore_() {
  // 🩹 [정밀검사 개선] 예전에는 백업이 실패해도 오류를 삼키고 복원을 계속 진행해, 화면의
  // "복원 직전 상태도 백업됩니다" 안내가 사실과 달라질 수 있었다. 이제 성공 여부를 반환하고,
  // 호출부(restoreBackup/importSnapshot/restoreLatestBackup_)는 실패하면 복원을 중단한다.
  try {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    runDailyBackup_();
    PropertiesService.getScriptProperties().setProperty(DESTRUCTIVE_SNAPSHOT_PROP_KEY_, todayStr);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// 🆕 [기능 추가] 대시보드에서 "마지막 백업이 언제였는지"를 바로 확인할 수 있도록,
// 백업 폴더에서 가장 최근 파일의 생성 시각을 찾아 반환한다. 지금까지는 이걸 확인하려면
// Apps Script 편집기에서 listBackupFiles_()를 수동 실행해야 했다. 트리거를 아직 설치하지
// 않았거나(createDailyBackupTrigger 미실행) 드라이브 접근 권한이 없으면 null을 반환하고,
// 이 경우 프런트에서 "백업 없음/트리거 설정 필요"로 안내한다.
function getLastBackupInfo_() {
  try {
    const folder = getOrCreateBackupFolder_();
    let latest = null;
    const it = folder.getFilesByType(MimeType.PLAIN_TEXT);
    while (it.hasNext()) {
      const f = it.next();
      if (f.getName().indexOf('백업_') !== 0) continue;
      const created = f.getDateCreated();
      if (!latest || created > latest) latest = created;
    }
    if (!latest) return { lastBackupAt: null, daysAgo: null };
    const daysAgo = Math.floor((new Date() - latest) / (1000 * 60 * 60 * 24));
    return { lastBackupAt: latest.toISOString(), daysAgo: daysAgo };
  } catch (e) {
    return { lastBackupAt: null, daysAgo: null };
  }
}

// 백업 폴더 안의 "백업_" 파일 중 최근 keepCount개만 남기고 나머지는 휴지통으로 이동.
function cleanupOldBackups_(folder, keepCount) {
  const files = [];
  const it = folder.getFilesByType(MimeType.PLAIN_TEXT);
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf('백업_') === 0) files.push(f);
  }
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  for (let i = keepCount; i < files.length; i++) {
    files[i].setTrashed(true);
  }
}

// ===================================================================
// ♻️ [기능 추가] 백업 복원
// ===================================================================
// runDailyBackup_()으로 만들어둔 JSON 스냅샷을 다시 시트로 되돌린다. 실수로 행/열을
// 대량 삭제했거나 잘못된 값으로 덮어썼을 때, 예전에는 백업 파일을 열어 값을 손으로
// 옮겨 적어야 했다. 이제 함수 실행 한 번으로 스냅샷 시점 상태로 되돌릴 수 있다.
//
// [사용 방법]
// 1) 이 Apps Script 편집기 상단 함수 목록에서 "listBackupFiles_" 선택 → ▶ 실행 →
//    실행 로그(보기 > 실행 로그)에서 원하는 백업 파일명/시각을 확인.
// 2) 최신 백업으로 되돌리려면 "restoreLatestBackup_"을 바로 실행.
//    특정 시점 백업으로 되돌리려면 "restoreFromBackupFile_" 함수 안의 fileId를
//    1번에서 확인한 파일 ID로 바꿔서 실행(또는 스크립트 편집기의 실행 > 함수 실행에서
//    직접 인자를 넘겨도 됨).
// ⚠️ 복원은 대상 시트의 기존 내용을 전부 지우고 스냅샷 값으로 덮어쓰며 되돌릴 수 없다.
//    걱정되면 복원 전에 runDailyBackup_()을 한 번 더 실행해 "복원 직전 상태"도
//    남겨두는 것을 권장한다.
// ===================================================================

function listBackupFiles_() {
  const folder = getOrCreateBackupFolder_();
  const files = [];
  const it = folder.getFilesByType(MimeType.PLAIN_TEXT);
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf('백업_') === 0) {
      files.push({ id: f.getId(), name: f.getName(), createdAt: f.getDateCreated() });
    }
  }
  files.sort((a, b) => b.createdAt - a.createdAt);
  files.forEach(f => Logger.log(`${f.name}  |  id: ${f.id}  |  생성: ${f.createdAt}`));
  return files;
}

/**
 * 지정한 백업 파일(fileId)의 스냅샷으로 스프레드시트의 모든 시트를 되돌린다.
 * 스냅샷에 없는 시트는 건드리지 않는다. 스냅샷에 있는 시트는 기존 내용을 전부 지우고
 * 스냅샷 값으로 다시 채운다(시트가 없으면 새로 생성).
 */
/**
 * 🩹 [정밀검사 개선] 백업 스냅샷은 저장 시점에 이미 날짜를 "yyyy-MM-dd" 문자열로
 * 정규화해뒀지만(runDailyBackup_), 복원 시 setValues로 그대로 쓰면 구글시트가 다시
 * 그 문자열을 Date/숫자로 자동 변환해버려, 방금 고친 날짜/티커 텍스트 서식 문제가
 * 복원 한 번으로 되살아날 수 있다. 시트 종류별로 텍스트 서식이 필요한 열(A열 등)에
 * 값을 쓰기 전에 미리 '@' 서식을 적용해 이를 막는다.
 */
function applyTextFormatForRestore_(sheet, sheetName, numRows, numCols) {
  if (numRows < 1) return;
  const dataRows = numRows - 1; // 헤더 제외
  if (sheetName === MAIN_SHEET_NAME) {
    // A열(날짜, 2행부터) + 티커 헤더 행(1행, I열부터)
    if (dataRows > 0) sheet.getRange(2, MAIN_COLS.date, dataRows, 1).setNumberFormat('@');
    if (numCols >= MAIN_COLS.tickerStart) {
      sheet.getRange(1, MAIN_COLS.tickerStart, 1, numCols - MAIN_COLS.tickerStart + 1).setNumberFormat('@');
    }
  } else if (sheetName === SHEET_NAME) {
    // A열(티커, 2행부터) + F열(종목목록, 2행부터)
    if (dataRows > 0) {
      sheet.getRange(2, 1, dataRows, 1).setNumberFormat('@');
      if (numCols >= 6) sheet.getRange(2, 6, dataRows, 1).setNumberFormat('@');
    }
  } else if (sheetName === CASHFLOW_SHEET) {
    if (dataRows > 0 && numCols >= 2) sheet.getRange(2, 2, dataRows, 1).setNumberFormat('@');
  } else if (sheetName === BUYMGMT_RECORDS_SHEET) {
    // B열(날짜, 2행부터)
    if (dataRows > 0 && numCols >= 2) sheet.getRange(2, 2, dataRows, 1).setNumberFormat('@');
  } else if (Object.values(BOOTSTRAP_GRID_SHEETS).indexOf(sheetName) >= 0) {
    // 그리드 시트: A열(행 라벨, 2행부터) + 헤더 행(1행, B열부터)
    if (dataRows > 0) sheet.getRange(2, 1, dataRows, 1).setNumberFormat('@');
    if (numCols >= 2) sheet.getRange(1, 2, 1, numCols - 1).setNumberFormat('@');
  }
}

// ===================================================================
// 🆕 [기능 추가 - JSON 백업 내보내기/불러오기] 스냅샷 검증 + 적용 공통 로직
// ===================================================================
// 대시보드에서 사용자가 직접 고른 JSON 백업 파일을 업로드해 복원할 수 있게 하려면,
// (드라이브 백업과 달리) 이 서버 코드 스스로가 신뢰할 수 없는 입력값을 받는다는
// 전제로 방어적으로 검증해야 한다. 아래 화이트리스트에 있는 시트만 복원 대상으로
// 허용하고, 셀 값도 문자열/숫자/불리언/빈 값만 허용한다.
function getImportSnapshotAllowedSheets_() {
  const set = {};
  set[SHEET_NAME] = 1;
  set[MAIN_SHEET_NAME] = 1;
  Object.values(BOOTSTRAP_GRID_SHEETS).forEach(name => { set[name] = 1; });
  set[BUYMGMT_RECORDS_SHEET] = 1;
  set[CASHFLOW_SHEET] = 1;
  return set;
}

function isAllowedSnapshotCellValue_(v) {
  if (v === null || v === undefined || v === '') return true;
  const t = typeof v;
  return t === 'string' || t === 'number' || t === 'boolean';
}

// 시트를 실제로 건드리기 전에 스냅샷 구조 자체가 올바른지 먼저 확인한다.
// 문제가 있으면 어떤 시트/행/열이 문제인지 사람이 이해할 수 있는 메시지를 반환한다.
function validateSnapshotStructure_(snapshot) {
  if (!snapshot || typeof snapshot !== 'object') return '스냅샷 형식이 올바르지 않습니다.';
  if (!snapshot.sheets || typeof snapshot.sheets !== 'object' || Array.isArray(snapshot.sheets)) {
    return '스냅샷에 sheets 데이터가 없습니다.';
  }
  const sheetNames = Object.keys(snapshot.sheets);
  if (sheetNames.length === 0) return '스냅샷에 시트 데이터가 없습니다.';

  for (let s = 0; s < sheetNames.length; s++) {
    const name = sheetNames[s];
    const values = snapshot.sheets[name];
    if (!Array.isArray(values)) return `'${name}' 시트 데이터가 2차원 배열이 아닙니다.`;
    for (let r = 0; r < values.length; r++) {
      if (!Array.isArray(values[r])) return `'${name}' 시트의 ${r + 1}번째 행이 배열이 아닙니다.`;
      for (let c = 0; c < values[r].length; c++) {
        if (!isAllowedSnapshotCellValue_(values[r][c])) {
          return `'${name}' 시트의 ${r + 1}행 ${c + 1}열 값이 허용되지 않는 형식입니다(문자열/숫자/불리언/빈 값만 허용).`;
        }
      }
    }
  }

  if (snapshot.formulas !== undefined && snapshot.formulas !== null) {
    if (typeof snapshot.formulas !== 'object' || Array.isArray(snapshot.formulas)) return 'formulas 형식이 올바르지 않습니다.';
    const fNames = Object.keys(snapshot.formulas);
    for (let i = 0; i < fNames.length; i++) {
      const list = snapshot.formulas[fNames[i]];
      if (!Array.isArray(list)) return `'${fNames[i]}' 시트의 수식 목록이 배열이 아닙니다.`;
      for (let j = 0; j < list.length; j++) {
        const it = list[j];
        if (!Array.isArray(it) || it.length !== 3 || !Number.isInteger(it[0]) || !Number.isInteger(it[1]) || typeof it[2] !== 'string') {
          return `'${fNames[i]}' 시트의 수식 항목 ${j + 1}번이 올바르지 않습니다.`;
        }
      }
    }
  }
  return null;
}

// 복원 시 외부 네트워크를 건드리는 함수(IMPORT*/IMAGE)가 든 수식은 적용하지 않는다.
const BLOCKED_FORMULA_RE_ = /\b(IMPORT[A-Z]*|IMAGE)\s*\(/i;
const MAX_RESTORE_FORMULAS_ = 5000;

/**
 * 스냅샷(snapshot.sheets)을 실제 시트에 덮어쓴다. restoreFromBackupFile_(드라이브
 * 자동 백업 복원)과 importSnapshot 액션(사용자가 올린 JSON 백업 파일 복원)이 이 함수를
 * 공유한다. 화이트리스트에 없는 시트명은 건드리지 않고 skipped 목록으로 알린다.
 */
function applySnapshot_(snapshot) {
  const structErr = validateSnapshotStructure_(snapshot);
  if (structErr) return { success: false, error: structErr };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allowedSheets = getImportSnapshotAllowedSheets_();
  const restoredSheets = [];
  const skipped = [];
  let formulasRestored = 0;
  let formulasSkipped = 0;

  Object.keys(snapshot.sheets || {}).forEach(sheetName => {
    if (!allowedSheets[sheetName]) { skipped.push(sheetName); return; }

    const values = snapshot.sheets[sheetName];

    // 🩹 [정밀검사 개선] 백업 시점에 비어 있던 시트는 현재 내용도 비운다(시트가 없으면 만들지 않음).
    if (!values || values.length === 0) {
      const existing = ss.getSheetByName(sheetName);
      if (existing) { existing.clearContents(); restoredSheets.push(sheetName); }
      return;
    }

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    const numRows = values.length;
    const numCols = values.reduce((max, row) => Math.max(max, row.length), 1);
    const padded = values.map(row => {
      const r = row.slice();
      while (r.length < numCols) r.push('');
      return r;
    });

    sheet.clearContents();
    applyTextFormatForRestore_(sheet, sheetName, numRows, numCols);

    // "="로 시작하는 문자열 값은 그대로 쓰면 구글시트가 수식으로 해석해버리므로,
    // 그 칸만 미리 텍스트 서식('@')을 적용해 문자 그대로 저장되게 한다.
    for (let r = 0; r < padded.length; r++) {
      for (let c = 0; c < padded[r].length; c++) {
        const v = padded[r][c];
        if (typeof v === 'string' && v.indexOf('=') === 0) {
          sheet.getRange(r + 1, c + 1).setNumberFormat('@');
        }
      }
    }

    sheet.getRange(1, 1, numRows, numCols).setValues(padded);
    restoredSheets.push(sheetName);

    // 🩹 [정밀검사 개선] 백업에 기록된 수식을 값 위에 다시 입힌다(범위 안 + 허용된 수식만).
    const fList = (snapshot.formulas && snapshot.formulas[sheetName]) || [];
    fList.forEach(item => {
      const r = item[0], c = item[1], f = item[2];
      if (formulasRestored >= MAX_RESTORE_FORMULAS_ || r < 1 || c < 1 || r > numRows || c > numCols ||
          f.indexOf('=') !== 0 || BLOCKED_FORMULA_RE_.test(f)) {
        formulasSkipped++;
        return;
      }
      try {
        const cell = sheet.getRange(r, c);
        cell.setNumberFormat('General');
        cell.setFormula(f);
        formulasRestored++;
      } catch (e) { formulasSkipped++; }
    });
  });

  // 시트1(일일 기록)이 복원됐다면 누적수익금/MDD/일별수익률 파생 컬럼을 처음부터 다시 계산한다.
  // (수식으로 복원된 칸이 있어도, 이 파생 컬럼은 스크립트가 값으로 관리하므로 그대로 재계산한다.)
  if (restoredSheets.indexOf(MAIN_SHEET_NAME) >= 0) {
    const mainSheet = getMainSheet_();
    if (mainSheet.getLastRow() >= 2) recalcDerivedFrom_(mainSheet, 2);
  }

  invalidateBootstrapCache_();

  return { success: true, restoredSheets: restoredSheets, skipped: skipped, backedUpAt: snapshot.backedUpAt || null,
           formulasRestored: formulasRestored, formulasSkipped: formulasSkipped };
}

function restoreFromBackupFile_(fileId) {
  if (!fileId) throw new Error('fileId가 필요합니다. listBackupFiles_()로 먼저 확인하세요.');
  const file = DriveApp.getFileById(fileId);
  const snapshot = JSON.parse(file.getBlob().getDataAsString());

  const result = applySnapshot_(snapshot);
  if (!result.success) throw new Error(result.error);

  Logger.log(`복원 완료 (백업 시각: ${result.backedUpAt}) - 시트: ${result.restoredSheets.join(', ')}` +
    (result.skipped.length > 0 ? ` / 건너뜀: ${result.skipped.join(', ')}` : ''));
  return { restoredSheets: result.restoredSheets, backedUpAt: result.backedUpAt };
}

// 가장 최근 백업 파일로 즉시 복원한다(편집기에서 함수 목록 선택 후 ▶ 실행).
function restoreLatestBackup_() {
  const files = listBackupFiles_();
  if (files.length === 0) {
    throw new Error('백업 파일이 없습니다. 먼저 runDailyBackup_()을 실행해 백업을 만들어주세요.');
  }
  // 🩹 [정밀검사 개선 - 문제 2] 대시보드 UI를 거치는 복원(doPost의 restoreBackup)은
  // snapshotBeforeRestore_를 호출하지만, 편집기에서 이 함수를 직접 실행하는 경로에는
  // 지금까지 그런 안전망이 전혀 없었다. 복원 경로가 무엇이든 동일하게 복원 직전
  // 상태를 남기도록 여기서도 호출한다.
  const pre = snapshotBeforeRestore_();
  if (!pre.ok) throw new Error('복원 직전 상태를 백업하지 못해 복원을 중단했습니다: ' + pre.error);
  return restoreFromBackupFile_(files[0].id);
}

// 매일 새벽 4시(스크립트 타임존 기준) 자동 백업 트리거를 설치한다. 최초 1회만 수동 실행.
function createDailyBackupTrigger() {
  deleteDailyBackupTrigger(); // 중복 생성 방지
  ScriptApp.newTrigger(BACKUP_TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();
  runDailyBackup_(); // 설치 확인용으로 즉시 1회 백업도 만들어둔다.
}

// 설치된 자동 백업 트리거를 전부 제거한다.
function deleteDailyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === BACKUP_TRIGGER_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

// ===================================================================
// 서버측 날짜 가드 / 입출금 기록 / 편집기 실행용 래퍼
// ===================================================================
const MAX_GAP_FILL_DAYS_ = 400;

function validateRecordDateRange_(sheet, dateStr) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (dateStr > addDays_(today, 1)) return '미래 날짜는 저장할 수 없습니다: ' + dateStr;
  const last = getLastDataRow_(sheet);
  if (last && last.dateStr && last.dateStr < dateStr) {
    const gap = Math.round((new Date(dateStr + 'T00:00:00') - new Date(last.dateStr + 'T00:00:00')) / 86400000) - 1;
    if (gap > MAX_GAP_FILL_DAYS_) return `마지막 기록(${last.dateStr})과 ${gap}일 떨어져 저장을 막았습니다(최대 ${MAX_GAP_FILL_DAYS_}일). 날짜(연도 오타)를 확인해주세요.`;
  }
  return null;
}

const CASHFLOW_SHEET = '입출금';
function getOrCreateCashflowSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(CASHFLOW_SHEET);
  if (!s) { s = ss.insertSheet(CASHFLOW_SHEET); s.getRange(1, 1, 1, 5).setValues([['ID', '날짜', '금액', '메모', '등록시각']]); }
  return s;
}
function readCashflows_(sheet) {
  const lastRow = sheet.getLastRow(); const out = [];
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 5).getValues().forEach(r => {
    const id = String(r[0] || '').trim(); if (!id) return;
    out.push({ id: id, date: normalizeCellText_(r[1]), amount: Number(r[2]) || 0, memo: String(r[3] || '') });
  });
  out.sort((a, b) => a.date < b.date ? -1 : (a.date > b.date ? 1 : 0));
  return out;
}
function addCashflow_(sheet, rec) {
  const id = 'c' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000);
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 2, 1, 1).setNumberFormat('@');
  sheet.getRange(row, 1, 1, 5).setValues([[id, rec.date, Number(rec.amount), String(rec.memo || ''), new Date()]]);
  return id;
}
function deleteCashflow_(sheet, id) {
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) { snapshotBeforeDestructiveAction_(); sheet.deleteRow(i + 2); return; }
  }
}

// 끝이 "_"인 함수는 편집기 실행 목록에 안 보일 수 있어, 실행용 래퍼를 둔다.
function setupApiToken() {
  const T = '여기에_토큰_입력'; // 실행 전 원하는 토큰 문자열로 바꾸세요(32자 이상 권장)
  if (T.indexOf('여기에') === 0) throw new Error('setupApiToken 함수 안의 T 값을 원하는 토큰으로 바꾼 뒤 실행하세요.');
  setApiToken_(T);
}
function runCleanupMainDates() { return cleanupMainSheetDuplicateDates_(); }
function runCleanupDuplicateHeaders() { return cleanupDuplicateTickerHeaders_(); }
function runBackupNow() { return runDailyBackup_(); }
function runRestoreLatestBackup() { return restoreLatestBackup_(); }

// ------------------- 응답 유틸 -------------------

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
