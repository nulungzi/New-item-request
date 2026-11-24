// ===== 소요제기 판정 웹앱 · Code.gs =====
// 시트·드라이브 입출력, 웹앱 함수(api*), 설정 점검(checkSetup). Apps Script 편집기의 Code.gs 에 이 파일 전체를 붙여넣습니다.
// Logic.gs 와 함께 있어야 동작합니다. Logic.gs 의 값(TPL 등)은 파일 맨 위가 아니라 함수 안에서만 씁니다 (파일 읽는 순서 때문).

/**
 * 소요제기 판정 웹앱 — 계약(LTA/3PP) 품목은 구매내역서를 만들고, Spot 품목은 견적 대상으로 분류한다.
 * 판정·계산은 1부(순수 함수), 여기 2부는 시트·드라이브 입출력만 맡는다.
 *
 * 스크립트 속성
 *   TEMPLATE_ID      (필수) 구매내역서 템플릿 구글 시트 ID — templates/구매내역서_템플릿.xlsx 를 드라이브에 올려 구글 시트로 저장
 *   OUTPUT_FOLDER_ID (권장) 생성한 구매내역서를 넣을 드라이브 폴더 ID
 *   SHEET_ID         (선택) 자재정보가 있는 스프레드시트 ID (없으면 이 스크립트가 붙은 시트)
 *   RFQ_TEMPLATE_ID  (견적요청서) 견적요청서 템플릿 구글 시트 ID — templates/견적요청서_템플릿.xlsx
 *   RFQ2_TEMPLATE_ID (견적요청서 2, 목록형) — templates/견적요청서_템플릿2_목록형.xlsx
 *   ARCHIVE_SHEET_ID (완료 보관) 완료된 기록을 옮겨 둘 다른 스프레드시트 ID
 *   SPOT_TEMPLATE_ID (Spot 견적) Spot 견적 구매내역서 템플릿 구글 시트 ID — templates/Spot_구매내역서_템플릿.xlsx
 */
var SHEET_MASTER = '자재정보';
var SHEET_RATE = '환율';
var SHEET_VENDORS = ['업체정보', '업체목록'];   // 업체명 · 이메일 열 (한 업체 여러 줄, 한 칸에 ; , 로 여러 주소 가능)
var SHEET_RECORD = '진행기록';     // 내가 만든 구매내역서·견적요청서 (완료 전)
var RECORD_HEADERS = ['ID', '종류', '담당자', '작성일', '제목', '파일 URL', '파일 ID', '상태', '메일 초안', '숨은 참조', '품목(JSON)', '수정일'];

/* ================================================================
 * 설정·시트 접근
 * ================================================================ */
function prop_(key) { return PropertiesService.getScriptProperties().getProperty(key); }

function db_() {
  var id = prop_('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function masterSheet_() {
  var sh = db_().getSheetByName(SHEET_MASTER);
  if (!sh) throw new Error('"' + SHEET_MASTER + '" 시트가 없습니다.');
  return sh;
}

/*
 * 자재정보 캐시: 판정·화면 열기 때마다 시트 전체를 읽지 않게 6시간 동안 기억 (모든 담당자 공용). 구매내역서 · 견적요청서를 만들 때는 새로 읽음.
 * 자재 등록 · DB 새로고침이면 새로 읽음. 날짜 칸은 'yyyy-mm-dd' 글자로 바꿔 저장 (시간대 밀림 없음)
 */
var CACHE_SEC = 600, CACHE_CHUNK = 30000, MASTER_CACHE_SEC = 21600;   // 자재정보는 6시간 (자재 등록 · DB 새로고침이면 바로 새로)
/** 공용 시트에 쓸 때 거는 잠금. 20초 넘게 못 받으면 안내 (데이터는 그대로) */
var LOCK_BUSY = '다른 분이 지금 저장하는 중이라 기다렸지만 끝나지 않았습니다. 잠시 후 다시 눌러 주세요.';
function lock_() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { throw new Error(LOCK_BUSY); }
  return lock;
}
function scriptCache_() { try { return CacheService.getScriptCache(); } catch (e) { return null; } }

/* ================================================================
 * 같은 요청 두 번 막기 · 이미 만든 뒤의 실패 안내
 *   메일 · 파일을 만든 다음 기록 저장이 실패하면, 예전에는 "아무것도 바뀌지 않았습니다" 로 끝나 다시 누르면
 *   메일이 두 번 나가거나 파일이 하나 더 생겼음. 이제:
 *   - once_: 같은 사람이 같은 내용으로 10분 안에 다시 누르면 다시 하지 않고 앞의 결과를 돌려줌 (repeated)
 *     처리하는 동안 또 누르면 "처리하는 중" 안내
 *   - afterEffect_: 이미 만든 뒤의 기록 저장은 한 번 더 해 보고, 그래도 안 되면 오류 대신 warn 으로 알림
 *   - trashOnFail_: 파일을 채우다 실패하면 반쯤 만든 파일은 휴지통으로
 * ================================================================ */
var ONCE_SEC = 600, ONCE_RUN_SEC = 360;
var ONCE_BUSY = '같은 요청을 지금 처리하고 있습니다. 끝날 때까지 기다려 주세요 (다시 누르지 않아도 됩니다).';
function userCache_() { try { return CacheService.getUserCache(); } catch (e) { return null; } }
/** 글자 → 짧은 키 (cyrb53 두 번) */
function hashKey_(s) {
  var h = function (seed) {
    var h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
    for (var i = 0; i < s.length; i++) { var ch = s.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (h2 >>> 0).toString(36) + (h1 >>> 0).toString(36);
  };
  return h(1) + h(7);
}
/**
 * stamp: 기록을 고치는 요청이면 그 기록의 지금 상태(수정일). 앞의 결과 뒤에 기록이 또 바뀌었으면
 *   (다른 내용으로 고친 뒤 처음 내용으로 다시 고치는 경우 등) 같은 내용이어도 새로 처리
 */
function once_(name, payload, fn, stamp) {
  var c = userCache_();
  if (!c) return fn();
  var key = 'once_' + name + '_' + hashKey_(JSON.stringify(payload === undefined ? null : payload)), hit = null;
  try { hit = JSON.parse(c.get(key) || 'null'); } catch (e) { hit = null; }
  if (hit && hit.state === 'run') throw new Error(ONCE_BUSY);
  if (hit && hit.state === 'done' && !(stamp && hit.after !== undefined && safeStamp_(stamp) !== hit.after)) {
    var r = hit.result && typeof hit.result === 'object' ? hit.result : {}; r.repeated = true; return r;
  }
  try { c.put(key, JSON.stringify({ state: 'run' }), ONCE_RUN_SEC); } catch (e) { /* 캐시가 안 되면 그냥 진행 */ }
  var result;
  try { result = fn(); } catch (e) { try { c.remove(key); } catch (e2) { /* 무시 */ } throw e; }
  try {
    var s = JSON.stringify({ state: 'done', result: result, after: stamp ? safeStamp_(stamp) : undefined });
    if (s.length < 90000) c.put(key, s, ONCE_SEC); else c.remove(key);
  } catch (e) { /* 무시 */ }
  return result;
}
function safeStamp_(stamp) { try { return String(stamp()); } catch (e) { return 'x'; } }
/** 기록의 지금 상태: 수정일 + 상태 (없으면 'none') */
function recStamp_(id) {
  var v = recordSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) {
    var u = v[i][11];
    return (Object.prototype.toString.call(u) === '[object Date]' ? u.getTime() : String(u)) + '|' + v[i][7];
  }
  return 'none';
}
/** what: '구매내역서 파일은 만들었지만' 처럼. 성공하면 '', 실패하면 화면에 보여 줄 안내 */
function afterEffect_(what, fn) {
  var err = null;
  for (var t = 0; t < 2; t++) {
    try { fn(); return ''; } catch (e) {
      err = e;
      if (String(e.message || e) === RECORD_TOO_BIG) break;   // 다시 해도 같음
      if (t === 0) { try { Utilities.sleep(1500); } catch (e2) { /* 시험 환경 */ } }
    }
  }
  console.warn(what + ' 기록 저장 실패: ' + err);
  return what + ' 진행 기록에는 남기지 못했습니다 (' + String((err && err.message) || err).replace(/\s*\(아무것도 바뀌지 않았습니다\)/, '') + '). ' +
    '다시 누르지 말고 새로고침해 확인해 주세요. 파일 · 메일은 그대로 있습니다.';
}
/** 새로 만든 스프레드시트를 채우다 실패하면 그 파일을 휴지통으로 옮기고 같은 오류를 냄 */
function trashOnFail_(ss, fn) {
  try { return fn(); } catch (e) {
    try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (e2) { console.warn('휴지통 이동 실패: ' + e2); }
    throw new Error(String(e.message || e) + ' — 만들던 파일은 드라이브 휴지통으로 옮겼습니다.');
  }
}
function trashFiles_(list) { (list || []).forEach(function (ss) { try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (e) { console.warn('휴지통 이동 실패: ' + e); } }); }

/* ================================================================
 * 변경기록 탭: 누가 · 언제 · 무엇을 (만들기 · 상태 · 견적 단가 · 업체 · 완료 · 삭제 · 되돌리기 · 업체정보 · 자재 등록 · 환율)
 *   한 줄씩 쌓기만 함. 탭은 경고 보호(고치려 하면 경고). 기록이 실패해도 본 작업은 계속
 * ================================================================ */
var SHEET_LOG = '변경기록', LOG_HEADERS = ['시각', '담당자', '동작', '기록 ID', '종류', '제목', '내용'];
function logSheet_() {
  var ss = db_(), sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold').setBackground('#c9daf8');
    sh.setFrozenRows(1);
    try {
      sh.getRange(1, 4, sh.getMaxRows(), 1).setNumberFormat('@');   // 기록 ID 는 글자로
      sh.protect().setDescription('변경기록 (사이트가 쌓음 · 고치지 마세요)').setWarningOnly(true);
    } catch (e) { console.warn('변경기록 서식·보호 실패: ' + e); }
  }
  return sh;
}
function logChange_(action, id, kind, title, detail) {
  try {
    var d = Array.isArray(detail) ? detail.join('\n') : String(detail == null ? '' : detail);
    logSheet_().appendRow([new Date(), me_(), action, String(id || ''), String(kind || ''), txt_(String(title || '')), txt_(d.slice(0, 5000))]);
  } catch (e) { console.warn('변경기록 실패: ' + e); }
}

/* ================================================================
 * 견적데이터 탭: 커진 견적(JSON)을 45,000자씩 나눠 담음 (기록 ID · 순서 · 내용 · 수정일)
 *   진행기록 · 완료기록에는 quoteRef 표시만 남고, 읽을 때 rowToRecord_ 가 여기서 붙여 줌
 * ================================================================ */
var SHEET_QDATA = '견적데이터', QDATA_CHUNK = 45000, QDATA_MARK = '~';   // 조각마다 앞에 ~ → = - + 로 시작해도 수식이 되지 않음 (읽을 때 뗌)
var qstoreMemo_ = null;   // 한 번 실행하는 동안 한 번만 읽음
function qdataSheet_(create) {
  var ss = db_(), sh = ss.getSheetByName(SHEET_QDATA);
  if (!sh && create) {
    sh = ss.insertSheet(SHEET_QDATA);
    sh.getRange(1, 1, 1, 4).setValues([['기록 ID', '순서', '견적(JSON 조각)', '수정일']]).setFontWeight('bold').setBackground('#c9daf8');
    sh.setFrozenRows(1);
    try {
      sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
      sh.getRange(1, 3, sh.getMaxRows(), 1).setNumberFormat('@');
      sh.protect().setDescription('견적데이터 (사이트가 씀 · 고치지 마세요)').setWarningOnly(true);
      sh.hideSheet();
    } catch (e) { console.warn('견적데이터 서식·보호 실패: ' + e); }
  }
  return sh;
}
function qstore_() {
  if (qstoreMemo_) return qstoreMemo_;
  var sh = qdataSheet_(false), parts = {}, out = {};
  if (sh) sh.getDataRange().getValues().slice(1).forEach(function (r) {
    var id = String(r[0] || '');
    var t = String(r[2] == null ? '' : r[2]);
    if (id) (parts[id] = parts[id] || []).push([Number(r[1]) || 0, t.charAt(0) === QDATA_MARK ? t.slice(1) : t]);
  });
  Object.keys(parts).forEach(function (id) { out[id] = parts[id].sort(function (a, b) { return a[0] - b[0]; }).map(function (p) { return p[1]; }).join(''); });
  qstoreMemo_ = out;
  return out;
}
/** fresh: 메모를 버리고 시트에서 새로 (잠금 안에서 고치기 전). 못 읽으면 null */
function loadQuoteOf_(id, fresh) {
  if (fresh) qstoreMemo_ = null;
  var s = qstore_()[String(id)];
  if (!s) return null;
  try { return JSON.parse(s); } catch (e) { console.warn('견적데이터 읽기 실패 ' + id + ': ' + e); return null; }
}
var QDATA_BROKEN = '이 기록의 견적(견적데이터 탭)을 읽지 못해 저장하지 않았습니다. 견적이 지워지지 않도록 멈췄습니다. 새로고침한 뒤 다시 해 보고, 계속되면 관리자에게 알려 주세요.';
/** 잠금 안에서 부름. quote 가 null 이면 그 기록의 줄을 지움 */
function writeQuoteOf_(id, quote) {
  var sh = qdataSheet_(true), text = quote ? JSON.stringify(quote) : '', now = new Date();
  var chunks = chunkText_(text, QDATA_CHUNK).map(function (c) { return QDATA_MARK + c; });
  var v = sh.getDataRange().getValues(), rows = [];
  for (var j = 1; j < v.length; j++) if (String(v[j][0]) === String(id)) rows.push(j + 1);
  if (rows.length === chunks.length) rows.forEach(function (r, k) { sh.getRange(r, 2, 1, 3).setValues([[k, chunks[k], now]]); });
  else {
    for (var d = rows.length - 1; d >= 0; d--) sh.deleteRow(rows[d]);
    if (chunks.length) sh.getRange(sh.getLastRow() + 1, 1, chunks.length, 4).setValues(chunks.map(function (c, k) { return [String(id), k, c, now]; }));
  }
  if (qstoreMemo_) { if (text) qstoreMemo_[String(id)] = text; else delete qstoreMemo_[String(id)]; }
}
function cachePut_(key, obj, sec) {
  var c = scriptCache_();
  if (!c) return;
  try {
    var parts = chunkText_(JSON.stringify(obj), CACHE_CHUNK), n = parts.length, m = {};
    for (var i = 0; i < n; i++) m[key + '_' + i] = parts[i];
    m[key + '_n'] = String(n);
    c.putAll(m, sec || CACHE_SEC);
  } catch (e) { console.warn('캐시 저장 실패: ' + e); }
}
function cacheGet_(key) {
  var c = scriptCache_();
  if (!c) return null;
  try {
    var n = Number(c.get(key + '_n'));
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var got = c.getAll(keys), s = '';
    for (var k = 0; k < keys.length; k++) { if (got[keys[k]] == null) return null; s += got[keys[k]]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function cacheDrop_(key) { var c = scriptCache_(); if (c) try { c.remove(key + '_n'); } catch (e) { /* 무시 */ } }
function masterValues_(fresh) {
  if (!fresh) { var hit = cacheGet_('master'); if (hit) return hit; }
  var values = masterSheet_().getDataRange().getValues().map(function (r) {
    return r.map(function (x) { return Object.prototype.toString.call(x) === '[object Date]' ? (isNaN(x) ? '' : fmtDate_(x)) : x; });
  });
  cachePut_('master', values, MASTER_CACHE_SEC);
  return values;
}
/* 자재정보 버전: 자재 등록 · DB 새로고침 때 바뀜 → 화면은 버전이 같으면 브라우저에 저장한 자재정보를 그대로 씀 */
function masterVer_() { return prop_('MASTER_VER') || '0'; }
function masterBump_() {
  var v = String(Date.now());
  try { PropertiesService.getScriptProperties().setProperty('MASTER_VER', v); } catch (e) { console.warn('버전 저장 실패: ' + e); }
  return v;
}
/** 자재 정보 화면: 머리글(열 키) + 값 (머리글 아래, Item 이 있는 줄). 화면이 버전과 함께 저장해 둠 */
function apiMasterList() {
  auth_();
  var values = masterValues_(), head = masterHeader_(values);
  var cols = head.headers.map(function (h) { return { key: keyForHeader_(h), header: String(h).replace(/\s+/g, ' ').trim() }; });
  var itemCol = cols.map(function (c) { return c.key; }).indexOf('item');
  var used = cols.map(function (c, i) { return c.key ? i : -1; }).filter(function (i) { return i >= 0; });
  var rows = values.slice(head.row).filter(function (r) { return String(r[itemCol] || '').trim(); })
    .map(function (r) { return used.map(function (i) { return r[i] === null || r[i] === undefined ? '' : r[i]; }); });
  return { ver: masterVer_(), cols: used.map(function (i) { return cols[i]; }), rows: rows, at: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') };
}

function loadContracts_() { return parseContracts_(masterValues_()); }

/** 자재정보 머리글 위치 {row(1부터), headers} */
function masterHeader_(values) {
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    if (values[r].some(function (h) { return keyForHeader_(h) === 'item'; })) return { row: r + 1, headers: values[r] };
  }
  throw new Error('자재정보 시트에서 "Item" 머리글을 찾지 못했습니다.');
}

/** 신규 등록 양식에 쓸 머리글 목록 (시트 열 순서) */
function masterFields_(values) {
  return masterHeader_(values).headers
    .map(function (h) { return { header: String(h).replace(/\s+/g, ' ').trim(), key: keyForHeader_(h) }; })
    .filter(function (f) { return f.key; });
}

/** 환율 시트: A열 통화, B열 환율(KRW), C열 수정일. 없으면 만든다. */
function rateSheet_() {
  var ss = db_(), sh = ss.getSheetByName(SHEET_RATE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RATE);
    sh.getRange(1, 1, 1, 3).setValues([['CUR', 'KRW 환율', '수정일']]);
  }
  return sh;
}

function loadRates_() { return rateTable_().rates; }
/**
 * 환율 탭 읽기: rates {USD: 1350}, info {USD: {at:'yyyy-mm-dd', auto: B열이 수식(예: GOOGLEFINANCE)인지}}
 * 수식 칸은 계산된 값을 쓰고, 사이트가 덮어쓰지 않음
 */
function rateTable_() {
  var rg = rateSheet_().getDataRange(), v = rg.getValues(), f = rg.getFormulas ? rg.getFormulas() : [];
  var rates = {}, info = {};
  v.slice(1).forEach(function (r, i) {
    if (!r[0]) return;
    var cur = String(r[0]).trim().toUpperCase(), fr = f[i + 1] || [];
    rates[cur] = r[1];
    var at = Object.prototype.toString.call(r[2]) === '[object Date]' ? fmtDate_(r[2]) : toDate_(r[2]) ? fmtDate_(toDate_(r[2])) : '';
    info[cur] = { at: at, auto: !!fr[1] };
  });
  return { rates: rates, info: info };
}

function saveRates_(rates) {
  var lock = lock_();
  try { saveRatesLocked_(rates); } finally { lock.releaseLock(); }
}
/**
 * 사이트에서 쓴 환율을 환율 탭에 남김. B열(환율)이 수식(GOOGLEFINANCE 등)인 통화는 건드리지 않고,
 * C열(수정일)도 수식이면 그대로 둠 — 수식이 값으로 바뀌어 풀리던 문제
 */
function saveRatesLocked_(rates) {
  var sh = rateSheet_(), rg = sh.getDataRange(), rows = rg.getValues(), fs = rg.getFormulas ? rg.getFormulas() : [], now = new Date(), changed = [];
  Object.keys(rates).forEach(function (cur) {
    var i = -1;
    rows.forEach(function (r, n) { if (i < 0 && n > 0 && String(r[0]).trim().toUpperCase() === cur) i = n; });
    if (i < 0) { rows.push([cur, rates[cur], now]); sh.appendRow([cur, rates[cur], now]); changed.push(cur + ' ' + rates[cur] + ' (새 통화)'); return; }
    var f = fs[i] || [];
    if (f[1]) return;                                   // 수식 환율: 시트는 그대로
    if (Number(rows[i][1]) !== Number(rates[cur])) changed.push(cur + ' ' + rows[i][1] + ' → ' + rates[cur]);
    if (f[2]) sh.getRange(i + 1, 2).setValue(rates[cur]);
    else sh.getRange(i + 1, 2, 1, 2).setValues([[rates[cur], now]]);
  });
  if (changed.length) logChange_('환율 저장', '', '환율', '', changed);
}

/** google.script.run 으로 보낼 수 있게 Date → 'yyyy-mm-dd' */
function plainDates_(o) {
  var out = {};
  Object.keys(o).forEach(function (k) { out[k] = Object.prototype.toString.call(o[k]) === '[object Date]' ? fmtDate_(o[k]) : o[k]; });
  return out;
}

/* ================================================================
 * 로그인
 *   웹앱은 "웹 앱에 액세스하는 사용자" 권한으로 실행해야 각자 드라이브·Gmail 임시보관함을 씀.
 *   로그인 = 담당자가 입력한 이메일이 지금 접속한 Google 계정과 같은지 확인 → 사용자 탭에 등록 →
 *   4시간 동안 유지 (사용자별 UserProperties 에 로그인 시각)
 * ================================================================ */
var SHEET_USERS = '사용자';
var USER_HEADERS = ['이메일', '처음 등록', '마지막 로그인'];
var DEPLOY_MSG = '웹앱이 "나(배포한 사람)" 권한으로 배포되어 있어 모든 파일·메일 초안이 배포한 사람 계정으로 만들어집니다.\n' +
  '배포한 사람이 Apps Script 편집기에서 배포 → 배포 관리 → 수정 → "다음 사용자 인증 정보로 실행: 웹 앱에 액세스하는 사용자" → 새 버전으로 배포해 주세요.';

function authState_() {
  var account = normEmail_(Session.getActiveUser().getEmail()), runAs = normEmail_(Session.getEffectiveUser().getEmail());
  var p = PropertiesService.getUserProperties();
  var at = Number(p.getProperty('LOGIN_AT')) || 0, who = normEmail_(p.getProperty('LOGIN_EMAIL'));
  var deployOk = !!account && account === runAs;
  var loggedIn = deployOk && who === account && sessionValid_(at, Date.now());
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { /* 편집기에서 실행 */ }
  return { account: account, deployOk: deployOk, deployMsg: deployOk ? '' : DEPLOY_MSG, loggedIn: loggedIn,
    expiresAt: loggedIn ? at + SESSION_HOURS * 3600000 : 0, now: Date.now(), hours: SESSION_HOURS, url: url };
}
/** 로그인했는지 확인하고 이메일을 돌려줌. 안 했거나 4시간이 지났으면 오류 (화면이 로그인 창을 띄움) */
function auth_() {
  var st = authState_();
  if (!st.deployOk) throw new Error('LOGIN_REQUIRED: ' + DEPLOY_MSG);
  if (!st.loggedIn) throw new Error('LOGIN_REQUIRED: 로그인해 주세요. (로그인은 ' + SESSION_HOURS + '시간 동안 유지됩니다)');
  return st.account;
}
function apiAuthState() { return authState_(); }

/** 담당자 이메일로 로그인. 처음이면 사용자 탭에 등록 */
function apiLogin(email) {
  var st = authState_(), typed = normEmail_(email);
  if (!st.deployOk) throw new Error(DEPLOY_MSG);
  if (!isEmail_(typed)) throw new Error('이메일 주소를 정확히 입력해 주세요.');
  if (typed !== st.account) {
    throw new Error('입력한 이메일(' + typed + ')과 지금 접속한 Google 계정(' + st.account + ')이 다릅니다.\n' +
      '본인 Google 계정으로 접속해야 본인 드라이브·Gmail 임시보관함에 만들어집니다. 아래 "다른 Google 계정으로 접속"을 눌러 계정을 바꿔 주세요.');
  }
  var lock = lock_();
  var isNew = false;
  try {
    var ss = db_(), sh = ss.getSheetByName(SHEET_USERS);
    if (!sh) { sh = ss.insertSheet(SHEET_USERS); sh.appendRow(USER_HEADERS); sh.setFrozenRows(1); }
    var now = new Date(), vals = sh.getDataRange().getValues(), row = 0;
    for (var i = 1; i < vals.length; i++) if (normEmail_(vals[i][0]) === typed) { row = i + 1; break; }
    if (row) sh.getRange(row, 3).setValue(now);
    else { sh.appendRow([typed, now, now]); isNew = true; }
  } finally { lock.releaseLock(); }
  PropertiesService.getUserProperties().setProperties({ LOGIN_AT: String(Date.now()), LOGIN_EMAIL: typed });
  var out = authState_();
  out.isNew = isNew;
  return out;
}
function apiLogout() {
  var p = PropertiesService.getUserProperties();
  p.deleteProperty('LOGIN_AT'); p.deleteProperty('LOGIN_EMAIL');
  return { ok: true };
}

/* ================================================================
 * 웹앱
 * ================================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('소요제기 판정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 사용법 팝업의 예시 사진·누를 곳 위치 (Guide.html, JSON). 파일이 없으면 빈 값 → 글만 보임 */
function apiGuide() {
  auth_();
  try { return HtmlService.createHtmlOutputFromFile('Guide').getContent(); } catch (e) { return '{"img":{},"hot":{}}'; }
}

/** 화면이 처음 열릴 때: 작성일, 저장된 환율, 구매일정 설정, Item 안내, 신규 등록 양식 */
function apiInit(fresh) {
  auth_();
  var values = masterValues_(!!fresh), rt;   // DB 새로고침이면 시트에서 새로
  return {
    masterVer: fresh ? masterBump_() : masterVer_(),
    today: fmtDate_(new Date()),
    rates: (rt = rateTable_()).rates,
    rateInfo: rt.info,   // 통화별 수정일 · 자동(수식) 여부 → 환율 칸 옆에 표시
    schedule: { tasks: SCHEDULE_TASKS, example: SCHEDULE_EXAMPLE, slots: SCHEDULE_SLOTS },
    items: itemIndex_(parseContracts_(values)),   // [Item, Description, Ship From, 계약 구분들, Shelf Life]
    fields: masterFields_(values)
  };
}

/** 소요제기 판정. input: 화면 표(행 배열) 또는 붙여넣기 글. baseDateStr: 작성일 (계약 판정은 품목별 Need-by 기준) */
function apiClassify(input, baseDateStr) {
  auth_();
  var parsed = Array.isArray(input) ? parseRows_(input) : parseRequest_(input);
  var base = toDate_(baseDateStr) || toDate_(new Date());
  var open = {};
  try { open = openItemIndex_(recordSheet_().getDataRange().getValues().slice(1).map(rowToRecord_).filter(function (r) { return r.id; }), me_()); } catch (e) { console.warn('진행기록 읽기 실패: ' + e); }
  var results = classify_(parsed.requests, loadContracts_(), base).map(function (r) {
    r.dups = (open[r.item] || []).slice(0, 3);   // 진행 중인 다른 소요제기에 같은 자재
    r.needBy = fmtDate_(r.needBy);
    r.candidates = r.candidates.map(plainDates_);
    r.alts = (r.alts || []).map(plainDates_);
    return r;
  });
  return { results: results, skipped: parsed.skipped, rates: loadRates_() };
}

/** 신규 자재를 자재정보 시트 맨 아래에 추가. rec 키는 masterFields_ 의 key. 머리글 없는 맨 앞 번호 열은 윗줄 + 1 */
function apiRegisterItem(rec) {
  auth_();
  var r = apiRegisterItems([rec]);
  if (!r.ok) throw new Error(r.errors[0].messages.join('\n'));
  var one = r.added[0];
  one.masterVer = r.masterVer;
  return one;
}

/**
 * 자재 여러 줄 등록 (자재 등록 화면·입력용 시트에서 붙여넣기). 모두 검사해서 하나라도 틀리면 아무것도 쓰지 않고
 * {ok:false, errors:[{i, item, messages}]}, 모두 맞으면 자재정보 맨 아래에 이어 붙이고 {ok:true, added:[...]}
 * 같은 묶음 안에서 겹치는 계약·SPOT 도 막음
 */
function apiRegisterItems(recs) {
  auth_();
  recs = recs || [];
  var filled = function (r) { return Object.keys(r || {}).some(function (k) { return String(r[k] == null ? '' : r[k]).trim() !== ''; }); };
  if (!recs.some(filled)) throw new Error('등록할 자재를 입력해 주세요.');
  var lock = lock_();
  try {
    var sh = masterSheet_(), values = sh.getDataRange().getValues(), head = masterHeader_(values);
    var known = parseContracts_(values), errors = [], ok = [];
    var typeKey = head.headers.map(keyForHeader_).filter(isTypeKey_)[0] || '';
    recs.forEach(function (rec, i) {   // i = 화면 표의 줄 번호(0부터), 빈 줄은 건너뜀
      if (!filled(rec)) return;
      var v = validateNewItem_(rec, known, typeKey);
      if (v.errors.length) { errors.push({ i: i, item: String(rec.item || ''), messages: v.errors }); return; }
      ok.push(v.rec);
      known.push({ item: v.rec.item, supplier: v.rec.supplier, priceMode: v.rec.priceMode, start: v.rec.start || null, end: v.rec.end || null, row: '이번 ' + (i + 1) + '번째 줄' });
    });
    if (errors.length) return { ok: false, errors: errors };

    var last = sh.getLastRow(), prev = values[last - 1] || [], num = !String(head.headers[0]).trim() && last > head.row && typeof prev[0] === 'number' ? prev[0] : null;
    var rows = ok.map(function (rec) { var row = newItemRow_(head.headers, rec); if (num !== null) row[0] = ++num; return row; });
    var n = head.headers.length, at = last + 1, target = sh.getRange(at, 1, rows.length, n);
    if (last > head.row) sh.getRange(last, 1, 1, n).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    head.headers.forEach(function (h, i) {
      var k = keyForHeader_(h);
      if (k === 'item') sh.getRange(at, i + 1, rows.length, 1).setNumberFormat('@');
      if (k === 'start' || k === 'end') sh.getRange(at, i + 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    });
    target.setValues(rows.map(function (r) { return r.map(txt_); }));   // = - + 로 시작하는 글자는 글자로
    logChange_('자재 등록', '', '자재정보', ok.length + '줄', ok.map(function (rec) { return rec.item + ' ' + (rec.priceMode || '') + ' ' + (rec.supplier || ''); }));
    cacheDrop_('master');   // 다음 판정은 새 자재 포함
    var ver = masterBump_();
    return { ok: true, masterVer: ver, added: ok.map(function (rec, i) { return { row: at + i, item: rec.item, desc: rec.desc || '', ship: rec.shipFrom || '', shelf: rec.shelfLife || '', mode: rec.priceMode }; }) };
  } finally {
    lock.releaseLock();
  }
}

/** 자재 입력용 시트: 머리글은 자재정보와 같은 열 순서 (복사해 사이트 자재 등록 표에 붙여넣기) */
function apiCreateItemSheet(folderInput) {
  auth_();
  var fields = masterFields_(masterValues_()), n = fields.length, ROWS = 300;
  if (!n) throw new Error('자재정보 머리글을 찾지 못했습니다.');
  var folder = targetFolder_(folderInput), ss = SpreadsheetApp.create('자재 등록 입력 시트');
  var sh = ss.getSheets()[0].setName('자재 등록');
  sh.getRange(1, 1, 1, n).setValues([fields.map(function (f) { return f.header; })])
    .setFontWeight('bold').setBackground('#c9daf8').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  var list = function (col, items) { sh.getRange(2, col, ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build()); };
  fields.forEach(function (f, i) {
    var c = i + 1, w = { item: 120, priceMode: 80, supplier: 110, desc: 320, product: 140, material: 110, qpl: 160, start: 95, end: 95 }[f.key] || 90;
    sh.setColumnWidth(c, w);
    sh.getRange(2, c, ROWS, 1).setNumberFormat(f.key === 'start' || f.key === 'end' ? 'yyyy-mm-dd' : f.key === 'price' ? '#,##0.####' : '@');
    if (f.key === 'priceMode') { list(c, ['SPOT', 'LTA', '3PP']); sh.getRange(1, c).setFontColor('#1c3d8f'); }
    if (f.key === 'item') sh.getRange(1, c).setFontColor('#1c3d8f');
    if (f.key === 'cur') list(c, ['USD', 'EUR', 'JPY', 'SGD', 'KRW']);
    if (f.key === 'shipFrom') list(c, ['Air', 'Ocean']);
  });
  sh.getRange(1, 1, ROWS + 1, n).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(1, 36); sh.setFrozenRows(1);
  if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);
  var g = ss.insertSheet('사용 방법');
  g.getRange(1, 1, 7, 1).setValues([['자재 등록 입력 시트 사용 방법'],
    ['1. "자재 등록" 탭 2행부터 한 줄에 한 자재씩 적습니다. 열 순서는 자재정보(DB)와 같으니 바꾸지 마세요.'],
    ['2. 꼭 필요한 칸: Item, 계약 구분(SPOT·LTA·3PP). LTA·3PP 는 업체명·단가·CUR·계약시작·계약종료도 필요합니다.'],
    ['3. 2행부터 마지막 줄까지 복사(Ctrl+C) → 웹앱 왼쪽 "자재 등록" 표의 첫 칸을 누르고 붙여넣기(Ctrl+V). 머리글을 같이 복사해도 됩니다.'],
    ['4. "자재정보에 등록" 을 누르면 모두 검사한 뒤 자재정보 맨 아래에 한 번에 추가합니다. 틀린 줄이 있으면 아무것도 쓰지 않고 그 줄을 알려 줍니다.'],
    ['5. 이미 있는 자재의 값을 고치는 것은 자재정보 시트에서 직접 합니다.'],
    ['날짜는 2027-12-31 처럼 적습니다.']]);
  g.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  g.setColumnWidth(1, 900);
  ss.setActiveSheet(sh);
  if (folder) DriveApp.getFileById(ss.getId()).moveTo(folder);
  return { url: ss.getUrl(), name: ss.getName(), folder: folder ? folder.getName() : '내 드라이브' };
}

/**
 * 구매내역서 생성.
 * payload = { baseDate, rates: {USD: 1339.3}, remarks: {업체명: '비고'},
 *             lines: [{item, qty, needBy, program, shelfLife, shipFrom, taxCode, contractRow}],
 *             schedule: [{start, len}] × 5 (구매일정, 반 달 단위),
 *             name: 구매내역서 제목 (비우면 자동), confirmer: 사업부 담당자 '홍길동 과장',
 *             folder: 저장할 드라이브 폴더 주소/ID (비우면 OUTPUT_FOLDER_ID 또는 기본 위치) }
 * 단가는 contractRow 로 자재정보를 다시 읽어 서버에서 확정한다(화면 값은 믿지 않음).
 */
function apiCreatePurchase(payload) {
  auth_();
  return once_('apiCreatePurchase', payload, function () { return apiCreatePurchase_(payload); });
}
function apiCreatePurchase_(payload) {
  var info = vendorInfoMap_();
  // 같은 자재·같은 계약 단가는 한 줄 (Need-by 여러 건은 Need-by 칸에 날짜(수량)으로)
  var lines = mergePurchaseLines_(purchaseLines_(payload)).map(function (l) {
    var v = vendorOf_(info, l.supplier);
    l.supplierName = v.supplier; l.site = v.site;
    return l;
  });
  var rates = checkedRates_(payload.rates, uniq_(lines.map(function (l) { return l.cur; })).filter(function (c) { return c !== 'KRW'; }));
  saveRates_(rates);
  var base = toDate_(payload.baseDate) || toDate_(new Date());
  // Spot 단가(DB) 품목은 따로 (계약과 같은 양식, 업체선정사유 Sole source)
  var groups = payload.split === false ? [{ spot: false, lines: lines }] : [
    { spot: false, lines: lines.filter(function (l) { return !isSpotMode_(l.priceMode); }) },
    { spot: true, lines: lines.filter(function (l) { return isSpotMode_(l.priceMode); }) }
  ].filter(function (g) { return g.lines.length; });
  var made = [], warns = [];
  var built = groups.map(function (g) {
    var sups = uniq_(g.lines.map(function (l) { return l.supplier; }));
    var name = String(payload.name || '').trim()
      ? purchaseFileName_(payload.name, base, sups) + (groups.length > 1 && g.spot ? ' (Spot)' : '')
      : purchaseFileName_('', base, sups).replace(' 구매내역서_', g.spot ? ' Spot 구매내역서_' : ' 구매내역서_');
    var curs = uniq_(g.lines.map(function (l) { return l.cur; })).filter(function (c) { return c !== 'KRW'; });
    var ss = null;
    try {
      ss = copyTemplate_(name, payload.folder);
      made.push(ss);
      applyPlan_(ss, buildPurchasePlan_(g.lines, curs, rates, base, payload.schedule, { title: name, confirmer: payload.confirmer }));
      var proto = ss.getSheetByName(TPL.PROTO);
      if (proto) ss.deleteSheet(proto);
      ss.setActiveSheet(ss.getSheetByName(TPL.SUMMARY));
    } catch (e) {
      trashFiles_(made);
      throw new Error(String(e.message || e) + (made.length ? ' — 만들던 파일은 드라이브 휴지통으로 옮겼습니다.' : ''));
    }
    return { g: g, ss: ss, name: name };
  });
  var files = built.map(function (b) {
    var g = b.g, ss = b.ss, name = b.name, rec = { id: '' };
    var warn = afterEffect_(name + ' 파일은 만들었지만', function () { rec = addRecord_({ kind: g.spot ? 'Spot 구매내역서' : '계약 구매내역서', title: name, url: ss.getUrl(), fileId: ss.getId(), status: '작성', group: payload.group || '',
      reqDates: payload.reqDates || [],
      items: g.lines.map(function (l) { return { item: l.item, qty: l.qty, needBy: fmtDate_(toDate_(l.needBy)), supplier: l.supplier, price: l.price, cur: l.cur,
        desc: l.desc || '', uom: l.uom || '', mfg: l.product || l.qpl || '', priceMode: l.priceMode || '', leadtime: l.leadtime || '', program: l.program || '', reqDate: l.reqDate || '',
        deadline: (poDeadline_(l.needs && l.needs.length ? l.needs[0].needBy : l.needBy, l.leadtime, l.shipFrom) || {}).date || '' }; }) }); });
    if (warn) warns.push(warn);
    return { url: ss.getUrl(), name: name, id: ss.getId(), recordId: rec.id, spot: g.spot };
  });
  return { files: files, url: files[0].url, name: files[0].name, id: files[0].id, recordId: files[0].recordId, warn: warns.join('\n') };
}

/** 화면에서 받은 줄 + 자재정보 계약 행 → 구매내역서 한 줄 */
function purchaseLines_(payload) {
  var byRow = {};
  // 파일을 만들 때는 캐시(최대 6시간) 대신 자재정보 시트를 새로 읽음 → 시트에서 단가·기간을 바로 고쳐도 반영
  parseContracts_(masterValues_(true)).forEach(function (c) { byRow[c.row] = c; });
  var lines = (payload.lines || []).map(function (l) {
    var c = byRow[l.contractRow];
    if (!c || c.item !== l.item) throw new Error(l.item + ': 자재정보가 바뀌었습니다. 다시 판정해 주세요.');
    if (c.price === null || c.price === undefined || c.price === '') throw new Error(l.item + ' (' + c.supplier + '): 자재정보에 단가가 비어 있습니다. 단가를 채운 뒤 다시 판정해 주세요.');
    var o = {};
    Object.keys(c).forEach(function (k) { o[k] = c[k]; });
    o.qty = Number(l.qty);
    o.program = l.program || '';                       // 비우면 시트에서 노란 칸
    if (l.shipFrom) o.shipFrom = l.shipFrom;           // 화면에 적으면 자재정보 값 대신
    if (l.shelfLife) o.shelfLife = l.shelfLife;
    if (l.taxCode !== undefined && l.taxCode !== '') o.taxCode = isNaN(+l.taxCode) ? l.taxCode : +l.taxCode;
    o.needBy = toDate_(l.needBy);
    o.remark = (payload.remarks || {})[c.supplier] || '';
    o.reqDate = l.reqDate || '';
    return o;
  });
  if (!lines.length) throw new Error('구매내역서에 넣을 계약 품목이 없습니다.');
  return lines;
}

/** 필요한 통화의 환율이 모두 양수인지 확인 */
function checkedRates_(input, curs) {
  var rates = {};
  Object.keys(input || {}).forEach(function (k) {
    var n = Number(input[k]);
    if (!(n > 0)) throw new Error(k + ' 환율을 입력해 주세요.');
    rates[k] = n;
  });
  curs.forEach(function (c) { if (!rates[c]) throw new Error(c + ' 환율을 입력해 주세요.'); });
  return rates;
}

/** 템플릿 시트를 복사해 새 스프레드시트를 연다. folderInput: 담당자가 고른 폴더 주소/ID (없으면 OUTPUT_FOLDER_ID) */
function copyTemplate_(name, folderInput) {
  var templateId = prop_('TEMPLATE_ID');
  if (!templateId) throw new Error('스크립트 속성 TEMPLATE_ID(구매내역서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  var folder = targetFolder_(folderInput), tpl = DriveApp.getFileById(templateId);
  var file = folder ? tpl.makeCopy(name, folder) : tpl.makeCopy(name);
  var ss = SpreadsheetApp.openById(file.getId());
  if (!contractTemplateCurrent_(ss)) {
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    throw new Error('구매내역서 템플릿이 예전 버전입니다 (구매세부내역 Need-by 옆 수량 열이 없음). ' +
      'templates/구매내역서_템플릿.xlsx 를 새로 올려 구글 시트로 저장하고 TEMPLATE_ID 를 바꿔 주세요.');
  }
  return ss;
}
/** 계약 구매내역서 템플릿이 T Need-by · U 수량 · V SUPPLIER · W SITE 가 있는 새 버전인지 (예시 머리글 7행) */
function contractTemplateCurrent_(ss) {
  var d = ss.getSheetByName(TPL.DETAIL);
  return !!d && String(d.getRange(7, 21).getValue()).trim() === '수량' && String(d.getRange(7, TPL.D_LAST).getValue()).trim() === 'SITE';
}

/** 저장할 드라이브 폴더: 화면에서 고른 폴더 → 스크립트 속성 OUTPUT_FOLDER_ID → null(기본 위치) */
function targetFolder_(folderInput) {
  var picked = parseFolderId_(folderInput);
  if (String(folderInput || '').trim() && !picked) throw new Error('드라이브 폴더 주소를 확인해 주세요. 폴더를 열었을 때 주소창의 주소를 그대로 붙여넣으면 됩니다.');
  var id = picked || prop_('OUTPUT_FOLDER_ID');
  if (!id) return null;
  try { var f = DriveApp.getFolderById(id); f.getName(); return f; }
  catch (e) { throw new Error('드라이브 폴더를 열 수 없습니다. 주소가 맞는지, 이 폴더에 파일을 만들 권한(편집자)이 있는지 확인해 주세요.'); }
}

/* ================================================================
 * 폴더 선택 창: 각자 내 드라이브의 폴더를 따라 들어가며 고른다
 * (웹앱 실행 사용자를 "웹 앱에 액세스하는 사용자"로 배포하면 각자 자기 드라이브가 보인다)
 * ================================================================ */
/** parentId 안의 폴더 목록. parentId 가 없으면 내 드라이브 맨 위 */
function apiListFolders(parentId) {
  auth_();
  var root = DriveApp.getRootFolder(), rootId = root.getId();
  var here;
  try { here = parentId ? DriveApp.getFolderById(parentId) : root; here.getName(); }
  catch (e) { throw new Error('폴더를 열 수 없습니다. 주소가 맞는지, 볼 수 있는 권한이 있는지 확인해 주세요.'); }
  var folders = [], it = here.getFolders();
  while (it.hasNext() && folders.length < 300) {
    var f = it.next();
    if (!f.isTrashed()) folders.push({ id: f.getId(), name: f.getName() });
  }
  folders.sort(function (a, b) { return a.name.localeCompare(b.name, 'ko'); });
  // 위치 표시: 내 드라이브 > … > 현재 폴더
  var path = [], cur = here, guard = 0;
  while (cur && guard++ < 15) {
    var isRoot = cur.getId() === rootId;
    path.unshift({ id: isRoot ? '' : cur.getId(), name: isRoot ? '내 드라이브' : cur.getName() });
    if (isRoot) break;
    var ps = cur.getParents();
    cur = ps.hasNext() ? ps.next() : null;
  }
  if (!path.length || path[0].name !== '내 드라이브') path.unshift({ id: '', name: '내 드라이브' });
  return { id: here.getId() === rootId ? '' : here.getId(), name: here.getId() === rootId ? '내 드라이브' : here.getName(), path: path, folders: folders };
}

/** 폴더 선택 창의 "새 폴더" */
function apiCreateFolder(parentId, name) {
  auth_();
  name = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  if (!name) throw new Error('폴더 이름을 입력해 주세요.');
  var parent = parentId ? DriveApp.getFolderById(parentId) : DriveApp.getRootFolder();
  var f = parent.createFolder(name);
  return { id: f.getId(), name: f.getName() };
}

/** 화면 "확인" 버튼: 폴더 이름을 돌려준다 */
function apiCheckFolder(folderInput) {
  auth_();
  var f = targetFolder_(folderInput);
  return f ? { name: f.getName(), url: f.getUrl(), picked: !!parseFolderId_(folderInput) } : { name: '내 드라이브 (기본 위치)', url: '', picked: false };
}

/* ================================================================
 * 입력용 시트: 엑셀처럼 적어 두고 웹앱 표에 복사·붙여넣기
 * ================================================================ */
function apiCreateInputSheet(folderInput, title) {
  auth_();
  var folder = targetFolder_(folderInput);
  var name = String(title || '').trim() || '소요제기 입력 시트';
  var ss = SpreadsheetApp.create(name);
  var cols = INPUT_SHEET.COLUMNS, n = cols.length, ROWS = 300;

  // 1) 입력 탭
  var sh = ss.getSheets()[0].setName(INPUT_SHEET.TAB);
  sh.getRange(1, 1, 1, n).setValues([cols.map(function (c) { return c[0]; })])
    .setFontWeight('bold').setBackground('#c9daf8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  cols.forEach(function (c, i) {
    sh.setColumnWidth(i + 1, c[5]);
    sh.getRange(2, i + 1, ROWS, 1).setNumberFormat(c[4]);
    sh.getRange(1, i + 1).setNote(c[1] + ' · ' + c[2] + '\n예: ' + c[3]);
    if (c[1] === '필수') sh.getRange(1, i + 1).setFontColor('#1c3d8f');
  });
  sh.getRange(2, 6, ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Air', 'Ocean'], true).setAllowInvalid(true).build());
  sh.getRange(2, 3, ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('날짜를 넣어 주세요. 예: 2027-06-10').build());
  sh.getRange(1, 1, ROWS + 1, n).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(1, 28);
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);

  // 2) 사용 방법 탭
  var g = ss.insertSheet(INPUT_SHEET.GUIDE);
  var guide = [
    ['소요제기 입력 시트 사용 방법', '', '', ''],
    ['', '', '', ''],
    ['1. 적기', '"' + INPUT_SHEET.TAB + '" 탭 2행부터 한 줄에 한 품목씩 적습니다. 열 순서를 바꾸지 마세요.', '', ''],
    ['2. 복사', '2행부터 마지막 품목까지 A~G열을 드래그해서 복사(Ctrl+C)합니다. 머리글(1행)을 같이 복사해도 괜찮습니다.', '', ''],
    ['3. 붙여넣기', '웹앱 "소요제기 입력" 표의 1번 Item 칸을 누르고 붙여넣기(Ctrl+V) 합니다. 줄이 모자라면 자동으로 늘어납니다.', '', ''],
    ['4. 다음에', '2행 아래 내용을 지우고 다시 씁니다. 이 시트는 계속 써도 됩니다.', '', ''],
    ['', '', '', ''],
    ['열', '필수', '설명', '예시']
  ].concat(cols.map(function (c) { return [c[0], c[1], c[2], c[3]]; }));
  g.getRange(1, 1, guide.length, 4).setValues(guide);
  g.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  g.getRange(3, 1, 4, 1).setFontWeight('bold').setFontColor('#1c3d8f');
  var head = 8;
  g.getRange(head, 1, 1, 4).setFontWeight('bold').setBackground('#c9daf8');
  g.getRange(head, 1, cols.length + 1, 4).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);
  g.getRange(head + 1, 4, cols.length, 1).setNumberFormat('@').setFontColor('#5a6680');
  g.setColumnWidth(1, 110); g.setColumnWidth(2, 60); g.setColumnWidth(3, 420); g.setColumnWidth(4, 140);
  g.getRange(3, 2, 4, 3).mergeAcross();
  g.getRange(1, 1, guide.length, 4).setVerticalAlignment('middle').setWrap(true);
  ss.setActiveSheet(sh);

  if (folder) DriveApp.getFileById(ss.getId()).moveTo(folder);
  return { url: ss.getUrl(), name: ss.getName(), folder: folder ? folder.getName() : '내 드라이브' };
}

/* ================================================================
 * 시트 쓰기: buildPurchasePlan_ 의 작업 목록 적용
 * ================================================================ */
/**
 * 속도: 바로 붙어 있는 같은 종류의 칸 쓰기(값 · 배경색 · 숫자 서식)가 같은 행의 이어진 열이면 한 번에 씀.
 *   순서는 그대로라 결과는 하나씩 쓸 때와 같음 (tests: 한 칸씩 쓴 결과와 칸마다 비교)
 */
var BATCHABLE = { set: 1, bg: 1, nf: 1 };
function batchOps_(ops) {
  var out = [];
  ops.forEach(function (o) {
    var last = out[out.length - 1];
    if (BATCHABLE[o.op] && !(o.op === 'bg' && !o.color) && last && last.op === o.op + 'Row' && last.s === o.s && last.r === o.r && last.c + last.vals.length === o.c) {
      last.vals.push(o.op === 'set' ? o.v : o.op === 'bg' ? o.color : o.f);
      return;
    }
    if (BATCHABLE[o.op] && !(o.op === 'bg' && !o.color)) out.push({ op: o.op + 'Row', s: o.s, r: o.r, c: o.c, vals: [o.op === 'set' ? o.v : o.op === 'bg' ? o.color : o.f] });
    else out.push(o);
  });
  return out;
}
/** 값 하나를 시트에 넣을 형태로: {date} → 시트 시간대의 날짜 */
function cellValue_(v, tz) {
  if (v && typeof v === 'object' && v.date) {
    var d = toDate_(v.date);
    if (d && tz && typeof Utilities !== 'undefined' && Utilities.parseDate) d = Utilities.parseDate(fmtDate_(d), tz, 'yyyy-MM-dd');
    return d;
  }
  return v;
}
function applyPlan_(ss, ops, opt) {
  var t0 = Date.now(), n0 = ops.length;
  if (!(opt && opt.batch === false)) ops = batchOps_(ops);
  var cache = {}, tz = ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : '';
  var sh = function (n) {
    if (!cache[n]) {
      cache[n] = ss.getSheetByName(n);
      if (!cache[n]) throw new Error('템플릿에 "' + n + '" 탭이 없습니다.');
    }
    return cache[n];
  };
  var FORMAT = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  var area = function (o) { return sh(o.s).getRange(o.r1, o.c1, o.r2 - o.r1 + 1, o.c2 - o.c1 + 1); };
  ops.forEach(function (o) {
    switch (o.op) {
      case 'fmt': sh(TPL.PROTO).getRange(o.from, o.c1, 1, o.c2 - o.c1 + 1).copyTo(sh(o.s).getRange(o.r, o.c1), FORMAT, false); break;
      case 'cellFmt': sh(o.fs).getRange(o.fr, o.fc).copyTo(sh(o.s).getRange(o.r, o.c), FORMAT, false); break;
      case 'rowFmt': sh(o.fs).getRange(o.fr, o.c1, 1, o.c2 - o.c1 + 1).copyTo(sh(o.s).getRange(o.r, o.c1), FORMAT, false); break;
      case 'set': setCell_(sh(o.s).getRange(o.r, o.c), o.v, tz); break;
      case 'setRow':
        if (o.vals.length === 1) setCell_(sh(o.s).getRange(o.r, o.c), o.vals[0], tz);
        else sh(o.s).getRange(o.r, o.c, 1, o.vals.length).setValues([o.vals.map(function (v) { return cellValue_(v, tz); })]);
        break;
      case 'bgRow': sh(o.s).getRange(o.r, o.c, 1, o.vals.length).setBackgrounds([o.vals]); break;
      case 'nfRow': sh(o.s).getRange(o.r, o.c, 1, o.vals.length).setNumberFormats([o.vals]); break;
      case 'nf': sh(o.s).getRange(o.r, o.c).setNumberFormat(o.f); break;
      case 'bg': sh(o.s).getRange(o.r, o.c).setBackground(o.color); break;
      case 'merge': area(o).merge(); break;
      case 'unmerge': area(o).breakApart(); break;
      case 'height': sh(o.s).setRowHeight(o.r, Math.round(o.pt * 4 / 3)); break;   // pt → px
      case 'bottom':
        sh(o.s).getRange(o.r, o.c1, 1, o.c2 - o.c1 + 1)
          .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        break;
      case 'list':
        sh(o.s).getRange(o.r1, o.c, o.r2 - o.r1 + 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(o.items, true).setAllowInvalid(true).build());
        break;
      case 'freeze': sh(o.s).setFrozenRows(o.rows); sh(o.s).setFrozenColumns(o.cols); break;
      case 'checkbox': sh(o.s).getRange(o.r1, o.c, o.r2 - o.r1 + 1, 1).insertCheckboxes(); break;
      case 'deleteCols': sh(o.s).deleteColumns(o.c, o.n); break;
      case 'ensureCols': { var ec = sh(o.s), have = ec.getMaxColumns(); if (have < o.n) ec.insertColumnsAfter(have, o.n - have); break; }
      case 'cf': {   // 조건부 서식 규칙을 맨 앞에 추가 (나중에 넣은 규칙이 우선)
        var cfSheet = sh(o.s), rules = cfSheet.getConditionalFormatRules();
        var rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(o.formula).setBackground(o.bg);
        if (o.font) rule = rule.setFontColor(o.font);
        rules.unshift(rule.setRanges([cfSheet.getRange(o.r1, o.c1, o.r2 - o.r1 + 1, o.c2 - o.c1 + 1)]).build());
        cfSheet.setConditionalFormatRules(rules);
        break;
      }
      case 'width': sh(o.s).setColumnWidth(o.c, o.px); break;
      case 'wrap': sh(o.s).getRange(o.r, o.c).setWrap(true); break;
      case 'clear': area(o).clear(); break;   // 값·서식 지움 (열 너비는 그대로)
      case 'border': {   // 표 테두리 (edge 'left' 면 왼쪽만). 검정 실선, style 'medium' 이면 굵게
        var bs = o.style === 'medium' ? SpreadsheetApp.BorderStyle.SOLID_MEDIUM : SpreadsheetApp.BorderStyle.SOLID;
        if (o.edge === 'left') area(o).setBorder(null, true, null, null, null, null, '#000000', bs);
        else area(o).setBorder(true, true, true, true, true, true, '#000000', bs);
        break;
      }
      default: throw new Error('알 수 없는 작업: ' + o.op);
    }
  });
  // 실행 로그(Apps Script 편집기 → 실행)에 걸린 시간: 속도 비교용
  if (typeof ScriptApp !== 'undefined') console.log('시트 채우기: 작업 ' + n0 + '개 → 시트 호출 ' + ops.length + '번, ' + (Date.now() - t0) + 'ms');
}

/** 값 넣기: {date:'yyyy-mm-dd'} 는 날짜, '='로 시작하면 수식 */
/**
 * 날짜는 시트의 시간대로 만든다. 스크립트 시간대(예: 서울)와 시트 시간대(예: 미국)가 다르면
 * new Date(2026, 10, 1) 이 시트에서 2026-10-31 로 보여 월 칸이 한 달씩 당겨지는 문제가 있었음
 */
function setCell_(cell, v, tz) {
  if (v && typeof v === 'object' && v.date) cell.setValue(cellValue_(v, tz));
  else if (typeof v === 'string' && v.charAt(0) === '=') cell.setFormula(v);
  else cell.setValue(v);
}

/* ================================================================
 * 미리보기: 만든 구매내역서 탭을 HTML로
 * 1순위 구글 시트 "웹페이지(zip)" 내보내기, 실패하면 값·배경·병합만 그린 간단한 표
 * ================================================================ */
function apiPreview(id) {
  auth_();
  SpreadsheetApp.flush();
  var exported = {};
  try { exported = exportHtml_(id); } catch (e) { console.warn('zip 내보내기 실패: ' + e); }
  var ss = SpreadsheetApp.openById(id);
  return [TPL.SUMMARY, TPL.DETAIL, TPL.SCHEDULE, TPL.CONFIRM].map(function (name) {
    if (exported[name]) return { name: name, html: exported[name], mode: 'export' };
    var sh = ss.getSheetByName(name);
    return { name: name, html: sh ? simpleHtml_(sh) : '<p>탭이 없습니다.</p>', mode: 'simple' };
  });
}

function exportHtml_(id) {
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=zip', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
  var css = '', pages = {};
  Utilities.unzip(resp.getBlob().setContentType('application/zip')).forEach(function (f) {
    var n = f.getName();
    if (/\.css$/i.test(n)) css += f.getDataAsString('UTF-8');
    else if (/\.html$/i.test(n)) pages[n.replace(/^.*\//, '').replace(/\.html$/i, '')] = f.getDataAsString('UTF-8');
  });
  var extra = 'body{margin:8px;background:#fff}.row-headers-background,.column-headers-background,' +
    '.row-header-wrapper,.freezebar-cell,.freezebar-origin-ltr{display:none!important}';
  Object.keys(pages).forEach(function (k) {
    pages[k] = pages[k].replace(/<link[^>]+sheet\.css[^>]*>/i, '').replace(/<\/head>/i, '<style>' + css + extra + '</style></head>');
  });
  return pages;
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

function simpleHtml_(sh) {
  var nr = Math.min(sh.getLastRow(), 80), nc = Math.min(sh.getLastColumn(), 30);
  if (!nr || !nc) return '<p>내용 없음</p>';
  var rg = sh.getRange(1, 1, nr, nc);
  var v = rg.getDisplayValues(), bg = rg.getBackgrounds(), fw = rg.getFontWeights(), fc = rg.getFontColors(),
    al = rg.getHorizontalAlignments(), fs = rg.getFontSizes(), wrap = rg.getWraps();
  var skip = {}, span = {};
  rg.getMergedRanges().forEach(function (m) {
    var r0 = m.getRow(), c0 = m.getColumn();
    if (r0 > nr || c0 > nc) return;
    span[r0 + ',' + c0] = [Math.min(m.getNumRows(), nr - r0 + 1), Math.min(m.getNumColumns(), nc - c0 + 1)];
    for (var i = 0; i < m.getNumRows(); i++) for (var j = 0; j < m.getNumColumns(); j++) if (i || j) skip[(r0 + i) + ',' + (c0 + j)] = 1;
  });
  var html = '<table style="border-collapse:collapse;font-family:Malgun Gothic,Arial,sans-serif;font-size:12px;table-layout:fixed"><colgroup>';
  for (var c = 1; c <= nc; c++) html += '<col style="width:' + sh.getColumnWidth(c) + 'px">';
  html += '</colgroup>';
  for (var r = 0; r < nr; r++) {
    html += '<tr style="height:' + sh.getRowHeight(r + 1) + 'px">';
    for (var k = 0; k < nc; k++) {
      var key = (r + 1) + ',' + (k + 1);
      if (skip[key]) continue;
      var sp = span[key], filled = v[r][k] !== '' || (bg[r][k] && bg[r][k] !== '#ffffff');
      var st = 'padding:2px 4px;vertical-align:middle;overflow:hidden;' + (wrap[r][k] ? 'white-space:pre-wrap;' : 'white-space:nowrap;') +
        'background:' + bg[r][k] + ';color:' + fc[r][k] + ';font-weight:' + fw[r][k] + ';font-size:' + fs[r][k] + 'pt;' +
        'text-align:' + (al[r][k] || 'general').replace('general-', '').replace('general', 'left') + ';' +
        (filled ? 'border:1px solid #bbb;' : '');
      html += '<td' + (sp ? ' rowspan="' + sp[0] + '" colspan="' + sp[1] + '"' : '') + ' style="' + st + '">' + esc_(v[r][k]) + '</td>';
    }
    html += '</tr>';
  }
  return html + '</table>';
}

/* ================================================================
 * 설정 점검 — 편집기에서 checkSetup 을 선택하고 ▶ 실행, 아래 "실행 로그"에서 결과 확인
 * ================================================================ */
function checkSetup() {
  var lines = [], ok = function (m) { lines.push('✅ ' + m); }, bad = function (m) { lines.push('❌ ' + m); };
  try {
    var values = masterValues_(), head = masterHeader_(values), contracts = parseContracts_(values);
    ok('자재정보: 머리글 ' + head.row + '행, 품목 ' + contracts.length + '줄');
    var keys = masterFields_(values).map(function (f) { return f.key; });
    ['item', 'priceMode', 'supplier', 'price', 'cur', 'start', 'end'].forEach(function (k) {
      if (keys.indexOf(k) < 0) bad('자재정보에 필수 열이 없습니다: ' + MASTER_COLUMNS[k][0]);
    });
  } catch (e) { bad('자재정보: ' + e.message); }

  var tid = prop_('TEMPLATE_ID');
  if (!tid) bad('스크립트 속성 TEMPLATE_ID 가 없습니다');
  else {
    try {
      var tpl = SpreadsheetApp.openById(tid);
      var missing = [TPL.SUMMARY, TPL.DETAIL, TPL.SCHEDULE, TPL.CONFIRM, TPL.PROTO].filter(function (n) { return !tpl.getSheetByName(n); });
      if (missing.length) bad('템플릿에 없는 탭: ' + missing.join(', ') + ' — templates/구매내역서_템플릿.xlsx 를 다시 올려 주세요');
      else if (!contractTemplateCurrent_(tpl)) bad('구매내역서 템플릿이 예전 버전입니다 (구매세부내역 Need-by 옆 수량 열 없음) — 새 템플릿을 올려 주세요');
      else ok('템플릿: ' + tpl.getName());
    } catch (e) { bad('템플릿을 열 수 없습니다 (ID·권한 확인): ' + e.message); }
  }
  var rid = prop_('RFQ_TEMPLATE_ID');
  if (!rid) bad('스크립트 속성 RFQ_TEMPLATE_ID 가 없습니다 (견적요청서 템플릿)');
  else { try { var rt = SpreadsheetApp.openById(rid); if (rt.getSheetByName(RFQ.SHEET)) ok('견적요청서 템플릿: ' + rt.getName()); else bad('견적요청서 템플릿에 "' + RFQ.SHEET + '" 탭이 없습니다'); } catch (e) { bad('견적요청서 템플릿을 열 수 없습니다: ' + e.message); } }
  var rid2 = prop_('RFQ2_TEMPLATE_ID');
  if (!rid2) bad('스크립트 속성 RFQ2_TEMPLATE_ID 가 없습니다 (견적요청서 템플릿 2 · 목록형)');
  else { try { var rt2 = SpreadsheetApp.openById(rid2); if (rt2.getSheetByName(RFQ2.SHEET)) ok('견적요청서 템플릿 2: ' + rt2.getName()); else bad('견적요청서 템플릿 2에 "' + RFQ2.SHEET + '" 탭이 없습니다'); } catch (e) { bad('견적요청서 템플릿 2를 열 수 없습니다: ' + e.message); } }
  var sid = prop_('SPOT_TEMPLATE_ID');
  if (!sid) bad('스크립트 속성 SPOT_TEMPLATE_ID 가 없습니다 (Spot 견적 구매내역서 템플릿)');
  else { try { var st = SpreadsheetApp.openById(sid), miss = [SPOT.SUMMARY, SPOT.DETAIL, SPOT.QUOTE, SPOT.PROTO].filter(function (n) { return !st.getSheetByName(n); });
    if (miss.length) bad('Spot 템플릿에 없는 탭: ' + miss.join(', '));
    else if (!spotTemplateCurrent_(st)) bad('Spot 템플릿이 예전 버전입니다 (선정 결과 Need-by 옆 수량 열 없음) — 새 템플릿을 올려 주세요');
    else ok('Spot 견적 구매내역서 템플릿: ' + st.getName()); } catch (e) { bad('Spot 템플릿을 열 수 없습니다: ' + e.message); } }
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) bad('스크립트 속성 ARCHIVE_SHEET_ID 가 없습니다 (완료 보관 파일)');
  else { try { ok('완료 보관 파일: ' + SpreadsheetApp.openById(aid).getName()); } catch (e) { bad('완료 보관 파일을 열 수 없습니다: ' + e.message); } }
  var vs = vendorSheet_();
  if (vs) ok(vs.getName() + ': ' + apiVendors().vendors.length + '개 업체'); else bad('자재정보 스프레드시트에 "업체정보" 탭이 없습니다 (업체명 · 이메일 열)');
  lines.push(typeof Gmail !== 'undefined' ? '✅ Gmail API 서비스: 사용 중 (Gmail 서명 자동 사용)' : 'ℹ️ Gmail API 서비스 없음 — 사이트에서 저장한 서명을 씁니다');
  var fid = prop_('OUTPUT_FOLDER_ID');
  if (fid) { try { ok('저장 폴더: ' + DriveApp.getFolderById(fid).getName()); } catch (e) { bad('OUTPUT_FOLDER_ID 폴더를 열 수 없습니다: ' + e.message); } }
  else lines.push('ℹ️ OUTPUT_FOLDER_ID 없음 — 구매내역서가 템플릿과 같은 위치에 저장됩니다');

  console.log(lines.join('\n'));
  return lines.join('\n');
}

/* ================================================================
 * 견적요청서 (RFQ) · Gmail 임시보관함 초안
 * ================================================================ */
/** 견적요청서 만들기. payload = {baseDate, title, due('yyyy-mm-dd'), folder, items:[{item, registered, name, desc, qpl, uom, remark, qty, needBy, program, shelfLife, shipFrom, taxCode}]} */
/**
 * payload.template: 'month' (기본, 템플릿 1 · 같은 자재 한 줄 · 월별 칸) | 'list' (템플릿 2 · Need-by 마다 한 줄)
 */
function apiCreateRfq(payload) {
  auth_();
  return once_('apiCreateRfq', payload, function () { return apiCreateRfq_(payload); });
}
function apiCreateRfq_(payload) {
  var list = payload.template === 'list';
  // 템플릿 1: 같은 자재는 Need-by 가 달라도 한 줄 · 템플릿 2: 같은 자재여도 Need-by 마다 다음 줄
  var rows = list ? splitRfqItems_(payload.items || []) : mergeRfqItems_(payload.items || []);
  var items = rows.map(function (it) {
    var o = {};
    Object.keys(it).forEach(function (k) { o[k] = it[k]; });
    o.needBy = toDate_(it.needBy);
    return o;
  });
  if (!items.length) throw new Error('견적요청서에 넣을 품목이 없습니다.');
  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var title = String(payload.title || '').trim() || (list ? rfq2FileName_(base, '', '', '') : rfqTitle_(base, '', '', undefined, payload.due));
  var key = list ? 'RFQ2_TEMPLATE_ID' : 'RFQ_TEMPLATE_ID', tid = prop_(key);
  if (!tid) throw new Error('스크립트 속성 ' + key + '(' + (list ? '견적요청서 템플릿 2 · 목록형' : '견적요청서 템플릿') + ' 시트 ID)를 설정해 주세요. README 참고.');
  var folder = targetFolder_(payload.folder), tpl = DriveApp.getFileById(tid);
  var file = folder ? tpl.makeCopy(title, folder) : tpl.makeCopy(title);
  var ss = SpreadsheetApp.openById(file.getId()), fixed = [];
  trashOnFail_(ss, function () {
    applyPlan_(ss, list ? buildRfq2Plan_(items) : buildRfqPlan_(items, base));
    // 하나로 만들기: 계약·Spot 단가 줄을 기록에 붙여 두었다가 견적 종합 때 같은 구매내역서에 넣음
    fixed = (payload.fixed || []).length ? fixedLinesFor_(payload.fixed) : [];
  });
  var rec = { id: '' };
  var warn = afterEffect_('견적요청서 파일은 만들었지만', function () {
    rec = addRecord_({ kind: '견적요청서', title: title, url: ss.getUrl(), fileId: ss.getId(), status: '견적요청서 작성', group: payload.group || '', fixed: fixed, reqDates: payload.reqDates || [],
      items: rows, due: payload.due || '', template: list ? 'list' : 'month' });
  });
  return { url: ss.getUrl(), name: title, id: ss.getId(), recordId: rec.id, warn: warn };
}

/**
 * 견적요청서 없이 견적 바로 입력: 업체 견적을 메일 없이 먼저 받았을 때.
 *   파일 · 메일 없이 기록만 만들고(견적요청서 기록과 같은 모양, direct), 업체를 고른 뒤 견적 입력 → 견적 종합 → Spot 견적 구매내역서
 *   payload = {title, items(견적요청서 만들기와 같은 품목), group, fixed, reqDates}
 */
function apiCreateDirectQuote(payload) {
  auth_();
  return once_('apiCreateDirectQuote', payload, function () { return apiCreateDirectQuote_(payload); });
}
function apiCreateDirectQuote_(payload) {
  payload = payload || {};
  var rows = mergeRfqItems_(payload.items || []);
  if (!rows.length) throw new Error('견적을 입력할 품목이 없습니다.');
  var title = String(payload.title || '').trim() || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd') + ' Material Spot (견적 직접 입력)';
  var fixed = (payload.fixed || []).length ? fixedLinesFor_(payload.fixed) : [];
  var rec = addRecord_({ kind: '견적요청서', title: title, url: '', fileId: '', status: '견적 직접 입력', group: payload.group || '', fixed: fixed, reqDates: payload.reqDates || [],
    items: rows, template: 'month', direct: true });
  return { recordId: rec.id, name: title };
}
/**
 * 견적을 입력할 업체 정하기 (메일 없이): 견적 직접 입력 기록, 또는 메일을 보낸 뒤 업체를 더하거나 뺄 때
 *   뺀 업체의 견적은 다음 저장 때 빠짐. 메일로 보낸 견적요청서는 상태를 그대로 둠
 */
function apiSetQuoteVendors(recordId, vendors) {
  auth_();
  var found = findRecord_(recordId, true);
  if (found.rec.kind !== '견적요청서') throw new Error('견적 기록이 아닙니다.');
  var names = uniq_((vendors || []).map(function (v) { return String(v || '').trim(); }));
  if (!names.length) throw new Error('견적을 입력할 업체를 한 곳 이상 골라 주세요.');
  var before = found.rec.vendors || [];
  var status = found.rec.direct && !found.rec.spotFile ? '견적 입력' : undefined;
  var added = names.filter(function (v) { return before.indexOf(v) < 0; }), removed = before.filter(function (v) { return names.indexOf(v) < 0; });
  updateRecord_(found, { status: status, data: { vendors: names }, log: '견적 업체 바꾸기', logDetail: (added.length ? ['더함: ' + added.join(', ')] : []).concat(removed.length ? ['뺌: ' + removed.join(', ')] : []) });
  return { ok: true, vendors: names, added: added, removed: removed };
}

/**
 * 견적요청서 수정: 메일을 보낸 뒤에도 품목 추가 · 제거 · 수량 · Need-by 를 바꿈.
 *   새 견적요청서 파일을 같은 폴더에 만들고(예전 파일은 '(수정 전 …)' 으로 이름만 바꿔 남김),
 *   입력한 견적은 같은 Item 끼리 옮김. 다음 메일 초안은 [Revised]
 *   payload = {recordId, items:[한 줄 = Need-by 한 건 {item, name, desc, qpl, uom, remark, qty, needBy, program}], baseDate}
 */
function apiReviseRfq(payload) {
  auth_();
  return once_('apiReviseRfq', payload, function () { return apiReviseRfq_(payload); }, function () { return recStamp_((payload || {}).recordId); });
}
function apiReviseRfq_(payload) {
  payload = payload || {};
  var found = findRecord_(payload.recordId, true), rec = found.rec;
  if (rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  if (rec.spotFile) throw new Error('이미 견적 구매내역서를 만든 견적요청서는 고칠 수 없습니다. 새 소요제기로 만들어 주세요.');
  var lines = (payload.items || []).filter(function (l) { return String(l.item || l.name || '').trim() && (Number(toNumber_(l.qty)) || 0) > 0; });
  if (!lines.length) throw new Error('견적요청서에 넣을 품목이 없습니다 (Item 과 수량을 적어 주세요).');
  var bad = lines.filter(function (l) { return !String(l.name || '').trim(); }).map(function (l) { return l.item; });
  if (bad.length && !rec.direct) throw new Error('Item name 을 적어 주세요: ' + bad.join(', '));
  var list = rfqTemplateOf_(rec) === 'list';
  var rows = list ? splitRfqItems_(lines) : mergeRfqItems_(lines);
  var stamp0 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  // 견적 옮기기는 잠금 안에서 지금 값으로 (파일을 만드는 사이 다른 창에서 입력한 견적도 같이 옮김)
  var out = {};
  var revise = function (stamp, extra) {
    return function (cur) {
      var m = remapRfqQuote_(cur.items || [], (cur.quote || {}).items, rows), q = {};
      Object.keys(cur.quote || {}).forEach(function (k) { q[k] = cur.quote[k]; });
      q.items = m.quoteItems; q.savedAt = new Date().toISOString();
      var rv = { no: ((cur.rev && cur.rev.no) || 0) + 1, at: stamp, added: m.added, removed: m.removed, before: (cur.items || []).length, after: rows.length };
      cur.items = rows; cur.quote = q; cur.rev = rv;
      Object.keys(extra || {}).forEach(function (k) { cur[k] = extra[k]; });
      out.rev = rv;
    };
  };
  if (rec.direct) {   // 견적 직접 입력: 파일이 없으니 품목과 견적만 고침
    updateRecord_(found, { edit: revise(stamp0), log: '견적요청서 수정 (직접 입력)', logDetail: ['품목 ' + (rec.items || []).length + ' → ' + rows.length + '줄'] });
    return { url: '', name: rec.title, rev: out.rev };
  }
  var items = rows.map(function (it) { var o = {}; Object.keys(it).forEach(function (k) { o[k] = it[k]; }); o.needBy = toDate_(it.needBy); return o; });
  var key = list ? 'RFQ2_TEMPLATE_ID' : 'RFQ_TEMPLATE_ID', tid = prop_(key);
  if (!tid) throw new Error('스크립트 속성 ' + key + ' 을 설정해 주세요. README 참고.');
  var old = DriveApp.getFileById(rec.fileId), parents = old.getParents(), folder = parents.hasNext() ? parents.next() : null;
  var file = folder ? DriveApp.getFileById(tid).makeCopy(rec.title, folder) : DriveApp.getFileById(tid).makeCopy(rec.title);
  var ss = SpreadsheetApp.openById(file.getId());
  trashOnFail_(ss, function () { applyPlan_(ss, list ? buildRfq2Plan_(items) : buildRfqPlan_(items, toDate_(payload.baseDate) || toDate_(new Date()))); });
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  var oldName = rec.title + ' (수정 전 ' + stamp.replace(/[-:]/g, '').replace(' ', '-').slice(2) + ')';
  try { old.setName(oldName); } catch (e) { console.warn('예전 견적요청서 이름 바꾸기 실패: ' + e); }
  var warn = afterEffect_('새 견적요청서 파일은 만들었지만', function () {
    updateRecord_(found, { url: ss.getUrl(), fileId: ss.getId(), status: '견적요청서 수정', edit: revise(stamp, { template: list ? 'list' : 'month' }),
      log: '견적요청서 수정', logDetail: ['품목 ' + (rec.items || []).length + ' → ' + rows.length + '줄', '새 파일 ' + ss.getUrl(), '예전 파일 이름: ' + oldName] });
  });
  var m0 = out.rev ? null : remapRfqQuote_(rec.items || [], (rec.quote || {}).items, rows);   // 기록에 못 남긴 경우 화면 표시용
  return { url: ss.getUrl(), name: rec.title, rev: out.rev || { no: ((rec.rev && rec.rev.no) || 0) + 1, at: stamp, added: m0.added, removed: m0.removed, before: (rec.items || []).length, after: rows.length }, warn: warn };
}

/** 하나로 만들기: 계약 · Spot 단가 줄을 기록에 붙일 모양으로 */
function fixedLinesFor_(lines) {
  return mergePurchaseLines_(purchaseLines_({ lines: lines })).map(function (l) {
    var o = {};
    ['item', 'row', 'supplier', 'cur', 'price', 'priceMode', 'material', 'desc', 'product', 'uom', 'qpl', 'leadtime', 'shipFrom', 'shelfLife', 'taxCode', 'program', 'qty', 'reqDate']
      .forEach(function (k) { if (l[k] !== undefined && l[k] !== null) o[k] = l[k]; });
    o.needBy = fmtDate_(toDate_(l.needBy));
    o.needs = (l.needs || []).map(function (x) { return { needBy: fmtDate_(toDate_(x.needBy)), qty: x.qty }; });
    return o;
  });
}

/**
 * 견적 입력용 구글 시트: 메일 초안에서 고른 업체마다 탭, Item code·이름·요청 수량·Need-by 를 채움.
 * 담당자가 업체 견적을 여기 적고 줄을 복사해 사이트 견적 입력 창에 붙여넣음. 이미 있으면 그 시트를 돌려줌
 */
function apiCreateQuoteSheet(recordId, folderInput) {
  auth_();
  return once_('apiCreateQuoteSheet', [recordId, folderInput], function () { return apiCreateQuoteSheet_(recordId, folderInput); }, function () { return recStamp_(recordId); });
}
function apiCreateQuoteSheet_(recordId, folderInput) {
  var found = findRecord_(recordId, true), rec = found.rec;
  if (rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  if (rec.quoteSheet) return { url: rec.quoteSheet, name: rec.title + ' 견적 입력' };
  if (!(rec.vendors || []).length) throw new Error('메일 초안을 먼저 만들어 업체를 골라 주세요.');
  var name = rec.title.replace(/^\[KE\] RFQ \/\/ /, '') + ' 견적 입력';
  var folder = targetFolder_(folderInput), ss = SpreadsheetApp.create(name);
  return trashOnFail_(ss, function () { return fillQuoteSheet_(found, rec, ss, name, folder); });
}
function fillQuoteSheet_(found, rec, ss, name, folder) {
  if (folder) DriveApp.getFileById(ss.getId()).moveTo(folder);
  var values = quoteSheetRows_(rec.items || []), n = values.length, W = QUOTE_SHEET.HEAD.length, F = QUOTE_SHEET.FIXED, used = {};
  var guide = ss.getSheets()[0].setName(sheetTabName_('사용 방법', used));
  guide.getRange(1, 1, 6, 1).setValues([[rec.title + ' — 견적 입력용 시트'],
    ['1. 업체 탭에 업체 견적서의 값을 F열(U/P)부터 적거나 붙여넣습니다. 회신이 없으면 비워 두고 사이트에서 "회신 없음"을 체크하세요.'],
    ['2. NO BID 는 U/P 칸에 NO BID 라고 적습니다. CUR 은 USD·EUR·JPY·SGD·KRW.'],
    ['3. 줄 전체(A~N)를 복사해 사이트 → 진행 중 → 그 업체 견적 입력 창 표에 붙여넣으면 Item code 로 맞춰 들어갑니다.'],
    ['4. 입력 칸(F~N)만 복사했다면 사이트 표에서 같은 칸을 누르고 붙여넣으세요.'],
    ['A~E열(No·Item code·Item name·요청 수량·Need-by)은 견적요청서 그대로이니 고치지 마세요.']]);
  guide.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  guide.setColumnWidth(1, 900);
  rec.vendors.forEach(function (v) {
    var sh = ss.insertSheet(sheetTabName_(v, used));
    sh.getRange(1, 1, n, W).setNumberFormat('@');
    sh.getRange(1, 1, n, W).setValues(values);   // 위에서 글자(@) 형식이라 = 로 시작해도 수식이 되지 않음
    sh.getRange(2, 1, Math.max(1, n - 1), 1).setNumberFormat('0');
    sh.getRange(1, 1, 1, W).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
    sh.getRange(1, 1, 1, F).setBackground('#d9d9d9');
    sh.getRange(1, F + 1, 1, W - F).setBackground('#fff2cc');
    if (n > 1) {
      sh.getRange(2, 1, n - 1, F).setBackground('#f3f3f3').setFontColor('#555555');
      sh.getRange(2, F + 2, n - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['USD', 'EUR', 'JPY', 'SGD', 'KRW'], true).setAllowInvalid(true).build());
      sh.getRange(2, W, n - 1, 1).setWrap(true);
    }
    sh.getRange(1, 1, n, W).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
    QUOTE_SHEET.WIDTHS.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setFrozenRows(1); sh.setFrozenColumns(F);
  });
  if (rec.vendors.length) ss.setActiveSheet(ss.getSheets()[1]);
  var warn = afterEffect_('견적 입력용 시트는 만들었지만', function () { updateRecord_(found, { data: { quoteSheet: ss.getUrl() }, log: '견적 입력용 시트', logDetail: [ss.getUrl()] }); });
  return { url: ss.getUrl(), name: name, warn: warn };
}

/** 업체정보 탭: [{name, emails:[...]}] */
function vendorSheet_() {
  var ss = db_();
  for (var i = 0; i < SHEET_VENDORS.length; i++) { var sh = ss.getSheetByName(SHEET_VENDORS[i]); if (sh) return sh; }
  return null;
}
/** Spot 템플릿이 선정 결과 뒤 W Need-by · X UOM · Y Tax Code · Z SUPPLIER · AA SITE 가 있는 새 버전인지 */
function spotTemplateCurrent_(ss) {
  var d = ss.getSheetByName(SPOT.DETAIL);
  var q = ss.getSheetByName(SPOT.QUOTE);
  return !!d && String(d.getRange(9, SPOT.D_NEED).getValue()).trim() === 'Need-by' && String(d.getRange(9, SPOT.D_QTY).getValue()).trim() === '수량' &&
    String(d.getRange(9, SPOT.D_TAX).getValue()).trim() === 'Tax Code' &&
    !!q && String(q.getRange(3, SPOT.Q_NBMODE).getValue()).trim() === 'Need-by 구분';
}
/** 업체정보 탭 SUPPLIER·SITE (업체명 기준). 탭이 없으면 빈 값 */
function vendorInfoMap_() {
  var sh = vendorSheet_();
  return sh ? vendorInfo_(sh.getDataRange().getValues()) : {};
}
/** 업체정보 탭에 연락처 추가(기존 업체) 또는 새 업체 등록. 탭이 없으면 만듦. 끝나면 새 업체 목록 */
function apiAddVendor(o) {
  auth_();
  var lock = lock_();
  try {
    var sh = vendorSheet_();
    if (!sh) { sh = db_().insertSheet(SHEET_VENDORS[0]); sh.appendRow(['업체명', '이메일', 'SUPPLIER', 'SITE', 'CONTACT']); sh.setFrozenRows(1); }
    var values = sh.getDataRange().getValues(), row = vendorRow_(values, o || {}), at = vendorInsertAt_(values, o || {});
    // 자리에 줄을 끼워 넣고 바로 위 줄 서식(테두리·가운데 정렬·숫자 형식)을 그대로 → 옆 칸으로 밀려 보이지 않음
    sh.insertRowAfter(at);
    var dest = sh.getRange(at + 1, 1, 1, row.length);
    if (at > vendorCols_(values).hr + 1) sh.getRange(at, 1, 1, row.length).copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    dest.setValues([row.map(txt_)]);
    logChange_(o && o.isNew ? '업체 등록' : '업체 연락처 추가', '', '업체정보', row[vendorCols_(values).cn] || '', String((o || {}).email || '') + (o && o.contact ? ' · ' + o.contact : ''));
  } finally { lock.releaseLock(); }
  return apiVendors();
}
function apiVendors() {
  auth_();
  var sh = vendorSheet_();
  if (!sh) return { vendors: [], missing: true };
  var values = sh.getDataRange().getValues(), info = vendorInfo_(values);
  return { vendors: parseVendors_(values).map(function (v) { var x = vendorOf_(info, v.name); v.supplier = x.supplier; v.site = x.site; return v; }), missing: false };
}

/** 메일 서명: Gmail 설정의 서명(고급 서비스 Gmail API 추가 시) → 없으면 사이트에 저장한 서명 */
function signature_() {
  try {
    if (typeof Gmail !== 'undefined') {
      var list = Gmail.Users.Settings.SendAs.list('me').sendAs || [];
      var me = list.filter(function (x) { return x.isPrimary; })[0] || list[0];
      if (me && me.signature) return { html: me.signature, source: 'gmail', gmailService: true };
    }
  } catch (e) { console.warn('Gmail 서명 읽기 실패: ' + e); }
  var saved = PropertiesService.getUserProperties().getProperty('SIGNATURE_HTML') || '';
  return { html: saved, source: saved ? 'saved' : 'none', gmailService: typeof Gmail !== 'undefined' };
}
function apiSignature() { auth_(); return signature_(); }
function apiSaveSignature(text) {
  auth_();
  var html = String(text || '').trim();
  if (html && !/<[a-z][\s\S]*>/i.test(html)) html = esc_(html).replace(/\n/g, '<br>');   // 글자로 적으면 줄바꿈만 살림
  PropertiesService.getUserProperties().setProperty('SIGNATURE_HTML', html);
  return signature_();
}

/** Gmail 임시보관함에 초안 만들기. 받는 사람은 비우고 고른 업체만 숨은 참조. payload = {recordId, bcc:[...]} */
/**
 * 메일 초안 (Gmail 임시보관함, 보내지 않음). payload = {recordId, bcc, vendors, remind?, due?}
 * remind: 같은 견적요청서의 재요청 — 제목 [Reminder], 재요청 문구, 고른 업체는 견적 입력 업체에 합쳐짐.
 * 보낸 기록은 data.mails 에 [{no, kind, at, vendors, count, due}] 로 쌓여 진행 중에 메일1 → 메일2 → 구매내역서 로 보임
 */
function apiCreateRfqDraft(payload) {
  auth_();
  return once_('apiCreateRfqDraft', payload, function () { return apiCreateRfqDraft_(payload); }, function () { return recStamp_((payload || {}).recordId); });
}
function apiCreateRfqDraft_(payload) {
  var bcc = uniq_((payload.bcc || []).map(function (m) { return String(m).trim(); }).filter(function (m) { return /@/.test(m); }));
  if (!bcc.length) throw new Error('메일을 받을 업체를 한 곳 이상 골라 주세요.');
  var found = findRecord_(payload.recordId, true);
  var rec = found.rec, remind = !!payload.remind;
  if (rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  if (rec.direct) throw new Error('견적요청서 없이 직접 입력한 견적이라 보낼 메일이 없습니다.');
  if (remind && !rec.draft) throw new Error('첫 견적 요청 메일 초안을 먼저 만들어 주세요.');
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + rec.fileId + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('견적요청서 파일을 엑셀로 내보내지 못했습니다 (HTTP ' + resp.getResponseCode() + ').');
  var blob = resp.getBlob().setName(rec.title + '.xlsx');
  var due = remind && toDate_(payload.due) ? fmtDate_(toDate_(payload.due)) : rec.due;
  var mails = (rec.mails || []).slice();
  if (!mails.length && rec.draft) mails.push({ no: 1, kind: '견적 요청', at: rec.created, vendors: rec.vendors || [], count: 0, due: rec.due });
  var first = mails[0];
  var revised = !remind && rfqRevisedPending_(rec);
  var mail = rfqEmail_(rec.items, dueText_(due), signature_().html, remind ? { sentOn: first ? dueText_(first.at) : '' } : revised ? { revised: true } : null);
  var subject = remind ? remindTitle_(rec.title, payload.due) : revised ? revisedTitle_(rec.title) : rec.title;
  var draft = GmailApp.createDraft('', subject, mail.text, { htmlBody: mail.html, bcc: bcc.join(','), attachments: [blob] });
  var names = uniq_((payload.vendors || []).map(function (v) { return String(v).trim(); }));
  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  mails.push({ no: mails.length + 1, kind: remind ? '재요청' : revised ? '수정 요청' : '견적 요청', at: now, vendors: names, count: bcc.length, due: due || '' });
  var data = { mails: mails };
  // 첫 요청은 고른 업체가 견적 입력 업체, 재요청은 기존 업체에 새로 고른 업체를 더함
  if (names.length) data.vendors = remind || revised ? uniq_((rec.vendors || []).concat(names)) : names;
  if (remind && due !== rec.due) data.due = due;
  var warn = afterEffect_('Gmail 임시보관함에 초안은 만들었지만', function () {
    updateRecord_(found, { status: remind ? '재요청 초안' : revised ? '수정 메일 초안' : '메일 초안', draft: draft.getId(), bcc: bcc.join(', '), data: data,
      log: remind ? '재요청 메일 초안' : revised ? '수정 메일 초안' : '메일 초안', logDetail: ['업체: ' + names.join(', '), '받는 주소 ' + bcc.length + '개 (숨은 참조)', '제목: ' + subject] });
  });
  return { draftId: draft.getId(), link: 'https://mail.google.com/mail/#drafts', bccCount: bcc.length, subject: subject, warn: warn };
}

/* ================================================================
 * 진행 기록 — 각자 자기 기록만. 완료하면 보관 파일(ARCHIVE_SHEET_ID)로 옮긴다
 * ================================================================ */
function me_() { return (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase(); }

function recordSheet_() {
  var ss = db_(), sh = ss.getSheetByName(SHEET_RECORD);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RECORD);
    sh.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]).setFontWeight('bold').setBackground('#c9daf8');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rowToRecord_(r) {
  var items = [];
  try { items = JSON.parse(r[10] || '[]'); } catch (e) {}
  var extra = items && !Array.isArray(items) ? items : null;
  return {
    id: String(r[0]), kind: r[1], owner: String(r[2]).toLowerCase(), created: Object.prototype.toString.call(r[3]) === '[object Date]' ? fmtDate_(r[3]) : String(r[3]),
    title: r[4], url: r[5], fileId: r[6], status: r[7], draft: r[8], bcc: r[9],
    items: extra ? extra.items : items, due: extra ? extra.due : '', vendors: extra && extra.vendors || [], quote: extra && (extra.quote || (extra.quoteRef ? loadQuoteOf_(r[0]) : null)) || null,
    group: extra && extra.group || '', parent: extra && extra.parent || '', fixed: extra && extra.fixed || [], quoteSheet: extra && extra.quoteSheet || '',
    spotFile: extra && extra.spotFile || '', reqDates: extra && extra.reqDates || [], po: extra && extra.po || null, mails: extra && extra.mails || [],
    confirmMail: extra && extra.confirmMail || null, rev: extra && extra.rev || null, template: extra && extra.template || '', direct: !!(extra && extra.direct),
    qfiles: extra && extra.qfiles || {}, qfolder: extra && extra.qfolder || '',
    updated: r[11] instanceof Date ? r[11].toISOString() : String(r[11] || '')
  };
}

function addRecord_(o) {
  var lock = lock_();
  try {
    var sh = recordSheet_(), now = new Date();
    var id = Utilities.formatDate(now, 'Asia/Seoul', 'yyMMdd-HHmmss') + '-' + (Math.random().toString(36) + '000000').slice(2, 8);   // 같은 초에 여러 개 만들어도 겹치지 않게
    var data = { items: o.items || [], due: o.due || '' };
    // group: 같은 소요제기에서 만든 파일 묶음, parent: 이 파일을 만든 견적요청서, fixed: 하나로 만들기 때 같이 들어갈 계약 줄
    ['group', 'parent', 'fixed', 'reqDates', 'template', 'direct'].forEach(function (k) { if (o[k] && (!Array.isArray(o[k]) || o[k].length)) data[k] = o[k]; });
    var json = JSON.stringify(data);
    if (json.length > RECORD_CELL_MAX) throw new Error(RECORD_TOO_BIG);
    sh.appendRow([id, o.kind, me_(), now, txt_(o.title), o.url, o.fileId, o.status, '', '', json, now]);
    logChange_('만들기', id, o.kind, o.title, o.url ? '파일 ' + o.url : '');
    return { id: id };
  } finally { lock.releaseLock(); }
}

/** 기록 찾기. mine 이면 다른 담당자 기록은 거부 */
function findRecord_(id, mine) {
  var sh = recordSheet_(), values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      var rec = rowToRecord_(values[i]);
      if (mine && rec.owner !== me_()) throw new Error('다른 담당자의 기록입니다.');
      return { sheet: sh, row: i + 1, values: values[i], rec: rec };
    }
  }
  throw new Error('기록을 찾지 못했습니다. 새로고침해 주세요.');
}

/**
 * 기록 고치기. 여러 사람이 동시에 써도 섞이지 않게: 잠금을 건 뒤 ID 로 줄을 다시 찾고(그 사이 다른 사람이
 * 완료해서 줄이 밀렸을 수 있음) 최신 품목(JSON)에 합쳐서 한 번에 씀
 */
function updateRecord_(found, patch) {
  var lock = lock_();
  try {
    var sh = recordSheet_(), values = sh.getDataRange().getValues(), id = String(found.values ? found.values[0] : found.rec.id), r = -1;
    for (var i = 1; i < values.length; i++) if (String(values[i][0]) === id) { r = i; break; }
    if (r < 0) throw new Error('기록을 찾지 못했습니다 (이미 완료되었을 수 있습니다). 새로고침해 주세요.');
    var row = values[r], json = row[10], detail = [];
    if (patch.data || patch.edit) {   // 품목(JSON) 칸에 값 합치기 (edit: 잠금 안에서 지금 값을 고침)
      var cur = {};
      try { cur = JSON.parse(row[10] || '{}'); } catch (e) {}
      if (Array.isArray(cur)) cur = { items: cur };
      var outside = !!cur.quoteRef;   // 견적을 견적데이터 탭에 따로 둔 기록
      if (outside) {
        cur.quote = loadQuoteOf_(id, true);   // 잠금 밖에서 읽어 둔 값이 아니라 지금 시트의 값
        if (!cur.quote) throw new Error(QDATA_BROKEN);
        delete cur.quoteRef;
      }
      if (patch.guard) patch.guard(cur);   // 잠금 안에서 지금 값으로 확인 (예: 다른 창에서 먼저 저장한 견적)
      var before = cur.quote, beforeText = outside ? JSON.stringify(before) : '';
      Object.keys(patch.data || {}).forEach(function (k) { cur[k] = patch.data[k]; });
      if (patch.edit) patch.edit(cur);
      if (patch.data && patch.data.quote) detail = quoteDiff_(before, cur.quote, cur.items);
      var pk = packRecord_(cur, outside);   // 시트 한 칸 한도: 커지면 견적을 견적데이터 탭으로
      if (pk.tooBig) throw new Error(RECORD_TOO_BIG);
      if (pk.out && !(outside && JSON.stringify(pk.quote) === beforeText)) writeQuoteOf_(id, pk.quote);   // 견적이 그대로면 다시 쓰지 않음
      json = pk.json;
    }
    if (patch.url !== undefined || patch.fileId !== undefined) sh.getRange(r + 1, 6, 1, 2).setValues([[patch.url !== undefined ? patch.url : row[5], patch.fileId !== undefined ? patch.fileId : row[6]]]);
    sh.getRange(r + 1, 8, 1, 5).setValues([[patch.status !== undefined ? patch.status : row[7], patch.draft !== undefined ? patch.draft : row[8],
      patch.bcc !== undefined ? patch.bcc : row[9], json, new Date()]]);
    // 변경기록: 견적 저장은 바뀐 값이 있을 때만
    var action = patch.log || (patch.status !== undefined && patch.status !== row[7] ? '상태: ' + patch.status : '수정');
    var extra = [].concat(patch.logDetail || []);
    if (patch.status !== undefined && patch.status !== row[7] && patch.log) extra.push('상태 ' + row[7] + ' → ' + patch.status);
    if (!patch.log && patch.data) extra.push('바꾼 값: ' + Object.keys(patch.data).join(', '));
    if (!(patch.log === '견적 저장' && !detail.length)) logChange_(action, id, row[1], row[4], extra.concat(detail));
  } finally { lock.releaseLock(); }
}

/** 내 진행 기록 (최근 순) */
function apiMyRecords() {
  auth_();
  var who = me_();
  var list = recordSheet_().getDataRange().getValues().slice(1).map(rowToRecord_)
    .filter(function (r) { return r.id && r.owner === who; });
  list.sort(function (a, b) { return String(b.id).localeCompare(String(a.id)); });
  var today = toDate_(new Date());
  list.forEach(function (r) { r.alert = recordAlert_(r, today); });
  return { me: who, records: list, archiveReady: !!prop_('ARCHIVE_SHEET_ID'), today: fmtDate_(today) };
}

/** 완료: 진행기록에서 보관 파일로 옮긴다 (본인 기록만) */
/* ================================================================
 * 구매 이력 (DB 구매이력 탭) — 완료할 때 올릴지 · 뺄 줄을 고름
 * ================================================================ */
function historySheet_() {
  var ss = db_(), sh = ss.getSheetByName(HISTORY.SHEET);
  if (!sh) {
    sh = ss.insertSheet(HISTORY.SHEET);
    sh.getRange(1, 1, 1, HISTORY.HEAD.length).setValues([HISTORY.HEAD]).setFontWeight('bold').setBackground('#c9daf8');
    sh.setFrozenRows(1);
  }
  return sh;
}
/** 완료할 기록들의 이력 후보 줄. 견적 구매내역서는 파일의 견적종합에서 선정(A열)한 줄을 읽음 */
function historyCandidates_(ids) {
  var mine = recordSheet_().getDataRange().getValues().slice(1).map(rowToRecord_).filter(function (r) { return r.id && r.owner === me_(); });
  var byId = {}; mine.forEach(function (r) { byId[r.id] = r; });
  var info = vendorInfoMap_(), rows = [], notes = [];
  var firstDate = function (r) { return String((r.reqDates || [])[0] || '').split(' ')[0]; };
  (ids || []).forEach(function (id) {
    var r = byId[id];
    if (!r) { notes.push(id + ': 기록을 찾지 못했습니다.'); return; }
    var push = function (h, k) {
      h.recId = r.id; h.url = r.url; h.key = r.id + '#' + k; h.title = r.title; h.kind = r.kind;
      if (!h.supplierName) h.supplierName = vendorOf_(info, h.supplier).supplier;
      rows.push(h);
    };
    if (r.kind === '견적요청서') return;   // 구매가 정해진 건 구매내역서
    if (r.kind === '견적 구매내역서') {
      var parent = r.parent && byId[r.parent], dates = {};
      if (parent) (parent.items || []).concat(parent.fixed || []).forEach(function (it) {
        var d = it.reqDate || ((it.parts || [])[0] || {}).reqDate;
        if (d) dates[itemKey_(it.item)] = d;
      });
      try {
        var q = SpreadsheetApp.openById(r.fileId).getSheetByName(SPOT.QUOTE);
        if (!q) throw new Error('견적종합 탭이 없습니다');
        quoteSheetSelected_(q.getDataRange().getValues()).forEach(function (h, k) {
          h.mode = /^lta$|^3pp$/i.test(h.reason) ? h.reason.toUpperCase() : 'SPOT 견적';
          h.reqDate = dates[h.item] || firstDate(r) || (parent ? firstDate(parent) : '');
          push(h, k);
        });
      } catch (e) { notes.push(r.title + ': 파일을 읽지 못했습니다 (' + e.message + ')'); }
      return;
    }
    (r.items || []).forEach(function (it, k) {   // 계약 · Spot 구매내역서
      var mode = String(it.priceMode || (r.kind === 'Spot 구매내역서' ? 'SPOT' : '')).toUpperCase();
      push({ item: itemKey_(it.item), desc: it.desc || '', supplier: it.supplier || '', price: it.price, cur: it.cur || '', qty: it.qty, uom: it.uom || '', mfg: it.mfg || '',
        lt: it.leadtime || '', mode: mode, reason: mode === 'SPOT' ? 'Sole source' : mode, program: it.program || '', reqDate: it.reqDate || firstDate(r) }, k);
    });
  });
  return { rows: rows, notes: notes };
}
function apiHistoryCandidates(ids) {
  auth_();
  var c = historyCandidates_(ids);
  c.rows.forEach(function (h) { if (h.price === undefined) h.price = ''; });
  return c;
}
/**
 * 완료 + 구매 이력. o = {ids, upload: true|false, skip: [후보 key]}
 * upload 이면 뺀 줄(skip)을 제외하고 구매이력 탭에 쌓은 뒤, 기록을 보관 파일로 옮김
 */
function apiCompleteGroup(o) {
  auth_();
  o = o || {};
  var ids = (o.ids || []).map(String), added = 0;
  if (!ids.length) throw new Error('완료할 기록이 없습니다.');
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) throw new Error('스크립트 속성 ARCHIVE_SHEET_ID(완료 보관 스프레드시트 ID)를 설정해 주세요.');
  // 파일을 읽는 느린 일(견적 구매내역서의 선정 줄)은 잠금 밖에서
  var rows = [];
  if (o.upload) {
    var skip = {}; (o.skip || []).forEach(function (k) { skip[k] = 1; });
    rows = historyCandidates_(ids).rows.filter(function (h) { return !skip[h.key]; });
  }
  var who = me_(), today = fmtDate_(new Date()), lock = lock_();
  try {
    // 한 번에: 모두 내 기록인지 먼저 확인 → 구매 이력(같은 기록 ID 줄은 바꿔 씀, 두 번 올라가지 않음) → 보관 파일로 이동
    ids.forEach(function (id) { findRecord_(id, true); });
    if (rows.length) {
      var sh = historySheet_();
      removeHistoryOf_(sh, ids, who);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, HISTORY.HEAD.length).setValues(rows.map(function (h) { return historyRow_(h, who, today); }));
      added = rows.length;
    }
    var arc = archiveSheet_(aid);
    ids.forEach(function (id) { completeLocked_(id, arc); });
  } finally { lock.releaseLock(); }
  return { ok: true, added: added, completed: ids.length };
}
/** 구매이력에서 이 기록들(기록 ID · 본인)로 올린 줄 지우기. 지운 줄 수 */
function removeHistoryOf_(hs, ids, who, keep) {
  if (!hs) return 0;
  var want = {}; ids.forEach(function (id) { want[String(id)] = 1; });
  var hv = hs.getDataRange().getValues(), idCol = HISTORY.HEAD.indexOf('기록 ID'), whoCol = HISTORY.HEAD.indexOf('담당자'), n = 0;
  for (var j = hv.length - 1; j >= 1; j--) {
    if (want[String(hv[j][idCol])] && String(hv[j][whoCol]).toLowerCase() === who) {
      if (keep) (keep[String(hv[j][idCol])] = keep[String(hv[j][idCol])] || []).unshift(hv[j].map(function (x) { return Object.prototype.toString.call(x) === '[object Date]' ? fmtDate_(x) : x; }));   // 삭제기록에 같이 보관
      hs.deleteRow(j + 1); n++;
    }
  }
  return n;
}
/** 자재 이력 조회: 자재정보 단가(계약·Spot) + 구매이력 (최근 순 100건) */
function apiItemHistory(item) {
  auth_();
  var k = itemKey_(item);
  if (!k) throw new Error('Item 을 입력해 주세요.');
  var contracts = loadContracts_().filter(function (c) { return c.item === k; }).map(plainDates_);
  var sh = db_().getSheetByName(HISTORY.SHEET), hist = [];
  if (sh) sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (itemKey_(r[2]) !== k) return;
    var o = {};
    HISTORY.HEAD.forEach(function (h, i) { var v = r[i]; o[h] = Object.prototype.toString.call(v) === '[object Date]' ? fmtDate_(v) : v; });
    hist.push(o);
  });
  hist.sort(function (a, b) { return String(b['등록일']).localeCompare(String(a['등록일'])); });
  return { item: k, contracts: contracts, history: hist.slice(0, 100), historyReady: !!sh };
}
/** 견적 종합 '직전 구매' 채우기: Item 별 가장 최근 구매 */
function apiLastPurchases(items) {
  auth_();
  var sh = db_().getSheetByName(HISTORY.SHEET);
  return sh ? lastPurchases_(sh.getDataRange().getValues(), items) : {};
}
/** 계약 만료 임박 (기본 90일) — 홈 화면 */
function apiExpiring(days) {
  auth_();
  return { days: days || 90, list: expiringContracts_(loadContracts_(), toDate_(new Date()), days || 90) };
}

/* ================================================================
 * 아침 요약 메일 — 각자 켜고 끔 (본인 계정의 시간 트리거, 평일 아침 본인에게만)
 * ================================================================ */
var DIGEST_HANDLER = 'dailyDigest';
function digestTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === DIGEST_HANDLER; });
}
function apiDigestState() {
  auth_();
  var up = PropertiesService.getUserProperties();
  return { on: digestTriggers_().length > 0, hour: Number(up.getProperty('DIGEST_HOUR')) || 8 };
}
function apiSetDigest(on, hour) {
  auth_();
  var h = Math.min(10, Math.max(6, Math.round(Number(hour) || 8)));
  digestTriggers_().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  var up = PropertiesService.getUserProperties();
  if (on) {
    ScriptApp.newTrigger(DIGEST_HANDLER).timeBased().everyDays(1).atHour(h).inTimezone('Asia/Seoul').create();
    up.setProperty('DIGEST_HOUR', String(h));
    try { up.setProperty('DIGEST_URL', ScriptApp.getService().getUrl() || ''); } catch (e) {}
  }
  return { on: !!on, hour: h };
}
/** 오늘 요약 (메일과 같은 내용) — 사이트에서 미리 보기 */
function apiDigestPreview() {
  auth_();
  return digestFor_(me_()) || { count: 0 };
}
function digestFor_(who) {
  var today = toDate_(new Date());
  var recs = recordSheet_().getDataRange().getValues().slice(1).map(rowToRecord_)
    .filter(function (r) { return r.id && r.owner === who; });
  recs.forEach(function (r) { r.alert = recordAlert_(r, today); });
  var url = PropertiesService.getUserProperties().getProperty('DIGEST_URL') || '';
  return digest_(recs, expiringContracts_(loadContracts_(), today, 30), today, url);
}
/** 시간 트리거: 평일만, 확인할 일이 있을 때만 본인에게 보냄 */
function dailyDigest() {
  var dow = Number(Utilities.formatDate(new Date(), 'Asia/Seoul', 'u'));   // 1=월 … 7=일
  if (dow > 5) return;
  var who = me_();
  if (!who) return;
  var d = digestFor_(who);
  if (!d) return;
  MailApp.sendEmail({ to: who, subject: d.subject, body: d.text, htmlBody: d.html, name: '소요제기' });
}

/* ================================================================
 * 사업부 확인 요청 메일 — 검토가 끝난 구매내역서 링크를 넣어 날인 요청
 *   mode 'draft' 는 Gmail 임시보관함, 'send' 는 바로 보냄 (화면에서 한 번 더 확인)
 * ================================================================ */
function splitEmails_(s) {
  return uniq_(String(s || '').split(/[,;\s]+/).map(normEmail_).filter(String));
}
function apiConfirmDefaults() {
  auth_();
  var up = PropertiesService.getUserProperties();
  return { to: up.getProperty('CONFIRM_TO') || '', cc: up.getProperty('CONFIRM_CC') || '', name: up.getProperty('CONFIRM_NAME') || '',
    body: up.getProperty('CONFIRM_BODY') || '', signature: signature_().html };
}
function apiConfirmMail(p) {
  auth_();
  return once_('apiConfirmMail', p, function () { return apiConfirmMail_(p); }, function () { return recStamp_((p || {}).recordId); });
}
function apiConfirmMail_(p) {
  p = p || {};
  var found = findRecord_(p.recordId, true), rec = found.rec;
  if (!/구매내역서/.test(rec.kind)) throw new Error('구매내역서 기록에서만 보낼 수 있습니다.');
  var to = splitEmails_(p.to), cc = splitEmails_(p.cc);
  if (!to.length) throw new Error('받는 사람 이메일을 적어 주세요.');
  var bad = to.concat(cc).filter(function (x) { return !isEmail_(x); });
  if (bad.length) throw new Error('이메일 형식이 아닙니다: ' + bad.join(', '));
  var subject = String(p.subject || '').trim() || rec.title;
  var mail = confirmMail_({ title: rec.title, url: rec.url, name: p.name, body: p.body }, signature_().html);
  var opt = { htmlBody: mail.html };
  if (cc.length) opt.cc = cc.join(',');
  var send = p.mode === 'send', link = 'https://mail.google.com/mail/#drafts';
  if (send) { GmailApp.sendEmail(to.join(','), subject, mail.text, opt); link = 'https://mail.google.com/mail/#sent'; }
  else GmailApp.createDraft(to.join(','), subject, mail.text, opt);
  // 여기부터는 메일을 이미 보냈거나(초안을 만들었거나) 한 뒤: 실패해도 오류로 끝내지 않음 (다시 누르면 두 번 나감)
  try {
    PropertiesService.getUserProperties().setProperties({ CONFIRM_TO: to.join(', '), CONFIRM_CC: cc.join(', '), CONFIRM_NAME: String(p.name || '').trim(), CONFIRM_BODY: String(p.body || '').trim() });
  } catch (e) { console.warn('사업부 확인 메일 기본값 저장 실패: ' + e); }
  var at = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  var log = { at: at, to: to, cc: cc, mode: send ? 'send' : 'draft', subject: subject };
  var warn = afterEffect_(send ? '메일은 보냈지만' : '초안은 만들었지만', function () {
    updateRecord_(found, { status: send ? '사업부 확인 요청' : '사업부 확인 초안', data: { confirmMail: log }, log: send ? '사업부 확인 메일 보냄' : '사업부 확인 메일 초안',
      logDetail: ['받는 사람: ' + to.join(', ') + (cc.length ? ' · 참조: ' + cc.join(', ') : ''), '제목: ' + subject] });
  });
  return { ok: true, sent: send, link: link, at: at, to: to, warn: warn };
}

/** 완료 보관 파일에서 내 기록 (최근 순, 최대 300건). 화면의 '완료 기록 보기' */
function apiArchivedRecords() {
  auth_();
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) return { records: [], ready: false };
  var sh = SpreadsheetApp.openById(aid).getSheetByName('완료기록');
  if (!sh) return { records: [], ready: true };
  var who = me_();
  var list = sh.getDataRange().getValues().slice(1).map(function (r) {
    var rec = rowToRecord_(r), done = r[RECORD_HEADERS.length];
    rec.done = Object.prototype.toString.call(done) === '[object Date]' ? fmtDate_(done) : String(done || '');
    return rec;
  }).filter(function (r) { return r.id && r.owner === who; });
  list.sort(function (a, b) { return String(b.done || b.id).localeCompare(String(a.done || a.id)); });
  return { records: list.slice(0, 300), ready: true };
}

/* ================================================================
 * 업체 성적표 · 견적 업체 추천 — 견적요청서 기록(진행 중 + 완료, 모든 담당자)과 구매이력에서
 *   화면에는 업체 이름과 숫자만 나감 (다른 담당자의 기록 내용은 보이지 않음)
 * ================================================================ */
function quoteRecordsAll_() {
  var hit = cacheGet_('qrecs');   // 10분 동안은 완료 보관 파일을 다시 열지 않음
  if (hit) return hit;
  var out = [];
  var slim = function (rec) {   // 성적·추천에 쓰는 값만 (가볍게 캐시)
    var q = rec.quote || {};
    return { id: rec.id, kind: rec.kind, created: String(rec.created || ''), vendors: rec.vendors || [],
      items: (rec.items || []).map(function (it) { return { item: it && it.item }; }),
      quote: { rates: q.rates || {}, noReply: q.noReply || {}, items: (q.items || []).map(function (x) {
        var qs = {};
        Object.keys((x && x.quotes) || {}).forEach(function (v) { var p = x.quotes[v] || {}; qs[v] = { price: p.price, cur: p.cur, lt: p.lt, sel: p.sel, noReply: p.noReply }; });
        return { quotes: qs, at: x && x.at, part: x && x.part };
      }) } };
  };
  var take = function (rows) { rows.forEach(function (r) { if (r[0] && r[1] === '견적요청서') out.push(slim(rowToRecord_(r))); }); };
  take(recordSheet_().getDataRange().getValues().slice(1));
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (aid) {
    try { var sh = SpreadsheetApp.openById(aid).getSheetByName('완료기록'); if (sh) take(sh.getDataRange().getValues().slice(1)); }
    catch (e) { console.warn('완료기록 읽기 실패: ' + e); }
  }
  cachePut_('qrecs', out);
  return out;
}
function historyValues_() {
  var hs = db_().getSheetByName(HISTORY.SHEET);
  return hs ? hs.getDataRange().getValues() : [];
}
function apiVendorStats() {
  auth_();
  return vendorStats_(quoteRecordsAll_(), historyValues_());
}
/** 이 견적요청서의 자재를 예전에 견적 냈거나 산 업체 (이 기록 자신은 빼고) */
function apiVendorHints(recordId) {
  auth_();
  var rec = findRecord_(recordId, true).rec;
  var items = (rec.items || []).map(function (it) { return it && it.item; }).filter(Boolean);
  if (!items.length) return [];
  var recs = quoteRecordsAll_().filter(function (r) { return r.id !== rec.id; });
  return vendorStats_(recs, historyValues_(), items).filter(function (s) { return s.quoted || s.bought; });
}

/**
 * 완료 취소: 완료기록에서 진행기록으로 되돌림 (본인 기록만). 상태는 단계·내용으로 다시 정함
 *   그 기록으로 올린 구매이력 줄은 지움 (다시 완료할 때 다시 올림)
 */
function restoredStatus_(rec) {
  if (rec.po && rec.po.stage) return rec.po.stage;
  if (rec.kind === '견적요청서') return rec.spotFile ? '견적 구매내역서 작성' : (rec.mails || []).length > 1 ? '재요청 초안' : rec.draft ? '메일 초안' : '작성';
  if (rec.confirmMail) return rec.confirmMail.mode === 'send' ? '사업부 확인 요청' : '사업부 확인 초안';
  return '작성';
}
function apiRestoreArchived(ids) {
  auth_();
  var want = {};
  (ids || []).forEach(function (id) { if (String(id || '').trim()) want[String(id)] = 1; });
  if (!Object.keys(want).length) throw new Error('되돌릴 기록을 골라 주세요.');
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) throw new Error('스크립트 속성 ARCHIVE_SHEET_ID(완료 보관 스프레드시트 ID)를 설정해 주세요.');
  var who = me_(), lock = lock_();
  try {
    var sh = SpreadsheetApp.openById(aid).getSheetByName('완료기록'), rs = recordSheet_(), restored = 0, skipped = 0, history = 0, back = {};
    if (sh) {
      var v = sh.getDataRange().getValues();
      for (var i = v.length - 1; i >= 1; i--) {
        if (!want[String(v[i][0])]) continue;
        if (String(v[i][2]).toLowerCase() !== who) { skipped++; continue; }
        var row = v[i].slice(0, RECORD_HEADERS.length);
        row[7] = restoredStatus_(rowToRecord_(row));
        row[11] = new Date();
        rs.appendRow(row.map(txt_));
        sh.deleteRow(i + 1); restored++; back[String(v[i][0])] = 1;
        logChange_('완료 취소', v[i][0], v[i][1], v[i][4], '진행 중으로 되돌림 · 상태 ' + row[7]);
      }
    }
    if (restored) history = removeHistoryOf_(db_().getSheetByName(HISTORY.SHEET), Object.keys(back), who);
    return { restored: restored, skipped: skipped, history: history };
  } finally { lock.releaseLock(); }
}

/**
 * 완료 기록 삭제 (본인 기록만). 드라이브 파일은 지우지 않음.
 *   바로 지우지 않고 보관 파일의 '삭제기록' 탭으로 옮김 → 화면의 '지운 기록'에서 되돌릴 수 있음
 *   withHistory: 그 기록으로 구매이력 탭에 올린 줄(기록 ID · 담당자 본인)도 빼서 같이 보관
 */
var TRASH_SHEET = '삭제기록', TRASH_EXTRA = ['완료일', '삭제일', '삭제한 사람', '구매이력(JSON)'];
function trashSheet_(aid) {
  var arc = SpreadsheetApp.openById(aid), sh = arc.getSheetByName(TRASH_SHEET);
  if (!sh) {
    sh = arc.insertSheet(TRASH_SHEET);
    sh.getRange(1, 1, 1, RECORD_HEADERS.length + TRASH_EXTRA.length).setValues([RECORD_HEADERS.concat(TRASH_EXTRA)]).setFontWeight('bold').setBackground('#f4cccc');
    sh.setFrozenRows(1);
  }
  return sh;
}
function apiDeleteArchived(ids, withHistory) {
  auth_();
  var want = {};
  (ids || []).forEach(function (id) { if (String(id || '').trim()) want[String(id)] = 1; });
  if (!Object.keys(want).length) throw new Error('지울 기록을 골라 주세요.');
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) throw new Error('스크립트 속성 ARCHIVE_SHEET_ID(완료 보관 스프레드시트 ID)를 설정해 주세요.');
  var who = me_(), lock = lock_();
  try {
    var deleted = 0, skipped = 0, history = 0, mine = [], now = new Date();
    var sh = SpreadsheetApp.openById(aid).getSheetByName('완료기록');
    if (!sh) return { deleted: 0, skipped: 0, history: 0 };
    var v = sh.getDataRange().getValues(), rows = [];
    for (var i = v.length - 1; i >= 1; i--) {
      if (!want[String(v[i][0])]) continue;
      if (String(v[i][2]).toLowerCase() !== who) { skipped++; continue; }
      rows.push({ i: i, row: v[i] }); mine.push(String(v[i][0]));
    }
    if (!rows.length) return { deleted: 0, skipped: skipped, history: 0 };
    var hist = {};
    if (withHistory) history = removeHistoryOf_(db_().getSheetByName(HISTORY.SHEET), mine, who, hist);
    var ts = trashSheet_(aid), W = RECORD_HEADERS.length;
    ts.getRange(ts.getLastRow() + 1, 1, rows.length, W + TRASH_EXTRA.length).setValues(rows.map(function (x) {
      var id = String(x.row[0]), h = hist[id] || [];
      return x.row.slice(0, W).map(txt_).concat([x.row[W] || '', now, who, h.length ? JSON.stringify(h) : '']);
    }));
    rows.forEach(function (x) {   // 아래에서부터 지워야 줄 번호가 안 밀림 (rows 는 아래부터)
      sh.deleteRow(x.i + 1); deleted++;
      logChange_('삭제 (삭제기록으로)', x.row[0], x.row[1], x.row[4], (hist[String(x.row[0])] || []).length ? '구매이력 ' + hist[String(x.row[0])].length + '줄도 같이' : '');
    });
    return { deleted: deleted, skipped: skipped, history: history };
  } finally { lock.releaseLock(); }
}
/** 지운 기록 (본인 것, 최근 순 200건) */
function apiDeletedRecords() {
  auth_();
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) return { records: [] };
  var sh = SpreadsheetApp.openById(aid).getSheetByName(TRASH_SHEET), who = me_(), W = RECORD_HEADERS.length;
  if (!sh) return { records: [] };
  var d = function (x) { return Object.prototype.toString.call(x) === '[object Date]' ? fmtDate_(x) : String(x || ''); };
  var list = sh.getDataRange().getValues().slice(1).filter(function (r) { return r[0] && String(r[2]).toLowerCase() === who; }).map(function (r) {
    var h = 0;
    try { h = (JSON.parse(r[W + 3] || '[]') || []).length; } catch (e) { h = 0; }
    return { id: String(r[0]), kind: r[1], title: r[4], created: d(r[3]), url: r[5], done: d(r[W]), deleted: d(r[W + 1]), history: h };
  });
  list.sort(function (a, b) { return String(b.deleted + b.id).localeCompare(String(a.deleted + a.id)); });
  return { records: list.slice(0, 200) };
}
/** 지운 기록 되돌리기 → 완료기록으로 (같이 뺀 구매이력 줄도 다시 올림) */
function apiRestoreDeleted(ids) {
  auth_();
  var want = {};
  (ids || []).forEach(function (id) { if (String(id || '').trim()) want[String(id)] = 1; });
  if (!Object.keys(want).length) throw new Error('되돌릴 기록을 골라 주세요.');
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) throw new Error('스크립트 속성 ARCHIVE_SHEET_ID(완료 보관 스프레드시트 ID)를 설정해 주세요.');
  var who = me_(), lock = lock_();
  try {
    var ts = SpreadsheetApp.openById(aid).getSheetByName(TRASH_SHEET), W = RECORD_HEADERS.length, restored = 0, history = 0;
    if (!ts) return { restored: 0, history: 0 };
    var arc = archiveSheet_(aid), v = ts.getDataRange().getValues(), hs = null;
    for (var i = v.length - 1; i >= 1; i--) {
      if (!want[String(v[i][0])] || String(v[i][2]).toLowerCase() !== who) continue;
      arc.appendRow(v[i].slice(0, W).map(txt_).concat([v[i][W] || new Date()]));
      var h = [];
      try { h = JSON.parse(v[i][W + 3] || '[]') || []; } catch (e) { h = []; }
      if (h.length) {
        hs = hs || historySheet_();
        hs.getRange(hs.getLastRow() + 1, 1, h.length, h[0].length).setValues(h.map(function (r) { return r.map(txt_); }));
        history += h.length;
      }
      ts.deleteRow(i + 1); restored++;
      logChange_('삭제 되돌리기', v[i][0], v[i][1], v[i][4], '완료기록으로' + (h.length ? ' · 구매이력 ' + h.length + '줄 다시 올림' : ''));
    }
    return { restored: restored, history: history };
  } finally { lock.releaseLock(); }
}

/**
 * PO 이후 상태 기록 (구매내역서 기록). po = {stage, approval, pr, po, poDate, memo}
 * 상태 칸에 단계가 보이고, 값은 기록의 품목(JSON)에 po 로 저장
 */
var PO_STAGES = ['작성', '품의 중', '품의 완료', 'PR 발행', 'PO 발행', '입고'];
function apiSetPoStatus(id, po) {
  auth_();
  po = po || {};
  var found = findRecord_(id, true);
  if (found.rec.kind === '견적요청서') throw new Error('견적요청서는 메일·견적 단계로 관리합니다. 구매내역서에서 상태를 기록하세요.');
  var stage = PO_STAGES.indexOf(po.stage) >= 0 ? po.stage : '작성';
  var clean = { stage: stage };
  ['approval', 'pr', 'po', 'poDate', 'memo'].forEach(function (k) { var v = String(po[k] == null ? '' : po[k]).trim().slice(0, 200); if (v) clean[k] = v; });
  if (clean.poDate && toDate_(clean.poDate)) clean.poDate = fmtDate_(toDate_(clean.poDate));
  updateRecord_(found, { status: stage, data: { po: clean }, log: 'PO 상태', logDetail: Object.keys(clean).map(function (k) { return k + ': ' + clean[k]; }) });
  return { ok: true, po: clean };
}

function apiCompleteRecord(id) {
  auth_();
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) throw new Error('스크립트 속성 ARCHIVE_SHEET_ID(완료 보관 스프레드시트 ID)를 설정해 주세요.');
  var lock = lock_();
  try {
    completeLocked_(id, archiveSheet_(aid));
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function archiveSheet_(aid) {
  var arc = SpreadsheetApp.openById(aid), sh = arc.getSheetByName('완료기록');
  if (!sh) {
    sh = arc.insertSheet('완료기록');
    sh.getRange(1, 1, 1, RECORD_HEADERS.length + 1).setValues([RECORD_HEADERS.concat(['완료일'])]).setFontWeight('bold').setBackground('#c9daf8');
    sh.setFrozenRows(1);
  }
  return sh;
}
/** 잠금 안에서: 진행기록 → 완료기록 (ID 로 다시 찾음) */
function completeLocked_(id, sh) {
  var found = findRecord_(id, true);
  var row = found.values.slice(0, RECORD_HEADERS.length);
  row[7] = '완료';
  sh.appendRow(row.map(txt_).concat([new Date()]));   // = 로 시작하는 제목이 수식이 되지 않게
  found.sheet.deleteRow(found.row);
  logChange_('완료', id, row[1], row[4], '완료기록으로 옮김');
}

/* ================================================================
 * Spot 견적 구매내역서 — 견적요청서 기록에서 견적값을 입력해 만든다
 * ================================================================ */
var SPOT_EXISTS = 'SPOT_EXISTS';
var QUOTE_CONFLICT = '이 견적은 다른 창(또는 다른 기기)에서 먼저 저장되어, 덮어쓰지 않도록 이번 저장은 하지 않았습니다. 새로고침한 뒤 이어서 입력해 주세요.';
/** 견적 입력 임시 저장. quote = {vendors, items:[{prev, program, quotes}], rates} */
function apiSaveQuotes(recordId, quote) {
  auth_();
  var found = findRecord_(recordId, true);
  if (found.rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  var base = String((quote || {}).base || ''), savedAt = new Date().toISOString();
  quote = quote || {};
  delete quote.base;
  quote.savedAt = savedAt;
  updateRecord_(found, { data: { quote: quote }, log: '견적 저장', guard: function (cur) {
    var was = cur.quote && cur.quote.savedAt;
    if (was && base !== String(was)) throw new Error(QUOTE_CONFLICT);
  } });
  return { ok: true, saved: savedAt, savedAt: savedAt };
}

/* ---------- 업체 견적서 파일 (PDF·엑셀 등) ----------
 * 견적 입력 창에서 올리면 견적요청서 파일이 있는 폴더(직접 입력 견적은 처음 고른 폴더)에 저장하고
 * 기록 data.qfiles = {업체: [{id, name, url, at}]} 에 링크를 남김 (견적 저장과 따로 → 서로 덮어쓰지 않음)
 */
var QFILE_MAX = 20 * 1024 * 1024;
function apiAttachQuoteFile(p) {
  auth_();
  p = p || {};
  var found = findRecord_(p.recordId, true), rec = found.rec, vendor = String(p.vendor || '').trim();
  if (rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  if (!vendor) throw new Error('업체를 골라 주세요.');
  var bytes = Utilities.base64Decode(String(p.data || ''));
  if (!bytes.length) throw new Error('빈 파일입니다.');
  if (bytes.length > QFILE_MAX) throw new Error('파일이 20MB 보다 큽니다. 드라이브에 직접 올려 주세요.');
  var folder = p.folder ? DriveApp.getFolderById(p.folder) : rec.qfolder ? DriveApp.getFolderById(rec.qfolder) : rec.fileId ? DriveApp.getFileById(rec.fileId).getParents().next() : null;
  if (!folder) throw new Error('견적서를 저장할 폴더를 골라 주세요.');
  var orig = String(p.name || '견적서').replace(/[\\/:*?"<>|]/g, '_');
  var name = vendor + '_견적서_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd') + '_' + orig;
  var file = folder.createFile(Utilities.newBlob(bytes, String(p.mime || 'application/octet-stream'), name));
  var entry = { id: file.getId(), name: name, url: file.getUrl(), at: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') }, out = {};
  try {
    updateRecord_(found, { log: '견적서 파일 붙임', logDetail: [vendor + ': ' + name, entry.url], edit: function (cur) {
      cur.qfiles = cur.qfiles || {};
      (cur.qfiles[vendor] = cur.qfiles[vendor] || []).push(entry);
      if (p.folder) cur.qfolder = p.folder;
      out = cur.qfiles;
    } });
  } catch (e) {
    try { file.setTrashed(true); } catch (e2) { /* 무시 */ }
    throw new Error(String(e.message || e) + ' — 올린 파일은 휴지통으로 옮겼습니다. 다시 붙여 주세요.');
  }
  return { file: entry, qfiles: out, qfolder: p.folder || rec.qfolder || '' };
}
/** 붙인 견적서 빼기: 링크를 지우고 파일은 드라이브 휴지통으로 (30일 안에 되살릴 수 있음) */
function apiRemoveQuoteFile(recordId, vendor, fileId) {
  auth_();
  var found = findRecord_(recordId, true), out = {}, hit = false;
  updateRecord_(found, { log: '견적서 파일 뺌', logDetail: [vendor + ': ' + fileId], edit: function (cur) {
    var list = (cur.qfiles || {})[vendor] || [];
    cur.qfiles = cur.qfiles || {};
    cur.qfiles[vendor] = list.filter(function (f) { if (f.id === fileId) { hit = true; return false; } return true; });
    if (!cur.qfiles[vendor].length) delete cur.qfiles[vendor];
    out = cur.qfiles;
  } });
  if (hit) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { console.warn('견적서 휴지통 이동 실패: ' + e); } }
  return { qfiles: out };
}

/**
 * 견적 종합에서: 이 견적요청서 품목이 진행 중인 다른 기록(모든 담당자)에도 있는지.
 * 같은 소요제기에서 나뉜 기록(같은 group)은 빼고, 품목 Item → [{title, owner, mine, kind, created}]
 */
function apiQuoteDups(recordId) {
  auth_();
  var all = recordSheet_().getDataRange().getValues().slice(1).map(rowToRecord_).filter(function (r) { return r.id; });
  var rec = all.filter(function (r) { return r.id === String(recordId); })[0];
  if (!rec) return { dups: {} };
  var others = all.filter(function (r) { return r.id !== rec.id && !(rec.group && r.group === rec.group); });
  return { dups: dupsFor_(rec, others, me_()) };
}

/**
 * payload = {recordId, baseDate, rates, folder, name, confirmer, schedule, vendors:[업체명],
 *            items:[{item, name, desc, material, qpl, uom, remark, qty, needBy, program, prev, quotes}]}
 */
function apiCreateSpotPurchase(payload) {
  auth_();
  return once_('apiCreateSpotPurchase', payload, function () { return apiCreateSpotPurchase_(payload); }, function () { return recStamp_((payload || {}).recordId); });
}
function apiCreateSpotPurchase_(payload) {
  var found = findRecord_(payload.recordId, true);
  if (found.rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  var tid = prop_('SPOT_TEMPLATE_ID');
  if (!tid) throw new Error('스크립트 속성 SPOT_TEMPLATE_ID(Spot 견적 구매내역서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  // 이미 만든 파일이 있으면 한 번 더 확인 (화면이 again 으로 다시 보냄) · 다른 창에서 견적을 먼저 저장했으면 멈춤
  if (found.rec.spotFile && !payload.again) throw new Error(SPOT_EXISTS + '|' + found.rec.spotFile + '|이 견적요청서로 이미 Spot 견적 구매내역서를 만들었습니다. 하나 더 만들려면 한 번 더 누르세요.');
  var base0 = payload.base === undefined ? null : String(payload.base || '');
  var quoteGuard = function (cur) { var was = cur.quote && cur.quote.savedAt; if (base0 !== null && was && String(was) !== base0) throw new Error(QUOTE_CONFLICT); };
  quoteGuard({ quote: found.rec.quote });
  var items = (payload.items || []).map(function (it) {
    var o = {};
    Object.keys(it).forEach(function (k) { o[k] = it[k]; });
    o.needBy = toDate_(it.needBy);
    return o;
  });
  // 하나로 만들기로 붙여 둔 계약 품목이 있으면 앞에 넣음 (견적종합에 계약 업체로 미리 선정)
  items = (found.rec.fixed || []).map(fixedSpotItem_).concat(items);
  if (!items.length) throw new Error('구매내역서에 넣을 품목이 없습니다.');
  var vendors = uniq_((payload.vendors || []).map(function (v) { return String(v).trim(); }));
  var rates = {};
  Object.keys(payload.rates || {}).forEach(function (k) { var n = Number(payload.rates[k]); if (n > 0) rates[k] = n; });
  var noRate = spotCurrencies_(items, vendors).filter(function (c) { return !(rates[c] > 0); });   // 환율이 없으면 원화가 0 이 되어 최저가로 보임
  if (noRate.length) throw new Error(noRate.join(' · ') + ' 환율을 입력해 주세요 (견적에 쓴 통화).');
  if (Object.keys(rates).length) saveRates_(rates);
  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var name = String(payload.name || '').trim() || found.rec.title.replace(/^\[KE\] RFQ \/\/ /, '') + ' 견적 구매내역서';
  var folder = targetFolder_(payload.folder), tpl = DriveApp.getFileById(tid);
  var file = folder ? tpl.makeCopy(name, folder) : tpl.makeCopy(name);
  var ss = SpreadsheetApp.openById(file.getId());
  if (!spotTemplateCurrent_(ss)) {
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    throw new Error('Spot 구매내역서 템플릿이 예전 버전입니다 (선정 결과 Need-by 옆 수량 열이 없음). ' +
      'templates/Spot_구매내역서_템플릿.xlsx 를 새로 올려 구글 시트로 저장하고 SPOT_TEMPLATE_ID 를 바꿔 주세요.');
  }
  trashOnFail_(ss, function () {
    applyPlan_(ss, buildSpotPlan_(items, vendors, rates, base, payload.schedule, { title: name, confirmer: payload.confirmer, vendorInfo: vendorInfoMap_() }));
    var proto = ss.getSheetByName(SPOT.PROTO);
    if (proto) ss.deleteSheet(proto);
    ss.setActiveSheet(ss.getSheetByName(SPOT.SUMMARY));
  });
  var savedAt = new Date().toISOString(), prevQuote = found.rec.quote || {}, quote = {};
  Object.keys(prevQuote).forEach(function (k) { quote[k] = prevQuote[k]; });   // 회신 없음 · 제목 · 담당자는 그대로 두고
  var prevItems = prevQuote.items || [];
  quote.vendors = vendors; quote.rates = payload.rates; quote.savedAt = savedAt;
  quote.items = (payload.items || []).map(function (it, i) {   // 나눈 줄(at · part)을 지켜야 다시 열었을 때 견적이 다른 품목으로 밀리지 않음
    var o = {}, pv = prevItems[i] || {};
    Object.keys(it).forEach(function (k) { o[k] = it[k]; });
    if (o.at === undefined && pv.at !== undefined) o.at = pv.at;
    if (o.part === undefined && pv.part !== undefined) o.part = pv.part;
    return o;
  });
  var rec = { id: '' };
  var warn = afterEffect_('Spot 견적 구매내역서 파일은 만들었지만', function () {
    updateRecord_(found, { status: '견적 구매내역서 작성', data: { quote: quote, spotFile: ss.getUrl() }, guard: quoteGuard, log: 'Spot 견적 구매내역서 만듦', logDetail: [ss.getUrl()] });
  });
  warn += (warn ? '\n' : '') + afterEffect_('Spot 견적 구매내역서 파일은 만들었지만', function () {
    if (rec.id) return;
    rec = addRecord_({ kind: '견적 구매내역서', title: name, url: ss.getUrl(), fileId: ss.getId(), status: '작성', parent: found.rec.id, group: found.rec.group || found.rec.id,
      items: items.map(function (it) { return { item: it.item, qty: it.qty, needBy: fmtDate_(it.needBy), supplier: it.fixed ? it.fixed.vendor : selectedVendor_(it), deadline: (spotDeadline_(it) || {}).date || '' }; }) });
  });
  return { url: ss.getUrl(), name: name, id: ss.getId(), recordId: rec.id, savedAt: savedAt, warn: warn.replace(/^\n+|\n+$/g, '') };
}
