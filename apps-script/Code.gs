/**
 * 종합 작업 도구 — 05 브랜드검색 노출 모니터링 (Google Apps Script)
 *
 * 역할
 *  1) 10분마다 등록된 키워드를 네이버(PC/모바일)에서 검색해 브랜드검색 광고가
 *     "우리 광고주 것으로" 노출되는지 확인하고 시트에 기록
 *  2) 하루 한 번(+수동) 네이버 검색광고 API로 브랜드검색 계약(time-contract) 기간 동기화
 *  3) 미노출 / 확인실패 / 계약 만료 임박 시 이메일 알림
 *  4) 웹앱(doGet/doPost)으로 index.html의 05 탭과 JSON으로 통신
 *
 * 설치 순서는 저장소의 `노출모니터링_설치안내.md` 참고.
 *
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성)에 넣어야 하는 값
 *   NAVER_API_KEY     네이버 검색광고 API 액세스 라이선스
 *   NAVER_API_SECRET  비밀키
 *   WEB_TOKEN         05 탭에서 입력할 접속 토큰 (아무 긴 문자열)
 *   ALERT_EMAIL       (선택) 알림 받을 이메일, 쉼표로 여러 개. 없으면 설정 시트 → 없으면 본인 계정
 */

var SHEET = { ADV: '광고주', KW: '키워드', LOG: '로그', CONTRACT: '계약', SET: '설정' };

var HEADERS = {};
HEADERS[SHEET.ADV] = ['광고주명', '네이버 customerId', '식별 문구(쉼표 구분)', '알림 이메일(선택)', '사용(Y/N)'];
HEADERS[SHEET.KW] = ['광고주명', '키워드', '매체(PC/M)', '사용(Y/N)', '최근 상태', '최근 확인시각', '연속 미노출', '연속 확인실패', '마지막 알림시각', '최근 근거'];
HEADERS[SHEET.LOG] = ['시각', '광고주명', '키워드', '매체', '결과', '근거', '메모'];
HEADERS[SHEET.CONTRACT] = ['광고주명', '광고그룹명', '계약명', '계약 상태', '노출 시작', '노출 종료', '남은 일수', '계약 ID', '갱신시각'];
HEADERS[SHEET.SET] = ['항목', '값', '설명'];

var DEFAULT_SETTINGS = [
  ['CHECK_INTERVAL_MIN', '10', '노출 확인 간격(분). 5/10/15/30 중 하나. 바꾼 뒤 setup() 또는 05 탭 "설정 저장"으로 트리거 재설치'],
  ['DOWN_ALERT_AFTER', '2', '연속 몇 번 미노출이면 알림을 보낼지 (10분 간격이므로 2 = 20분 연속)'],
  ['FAIL_ALERT_AFTER', '3', '연속 몇 번 확인실패(차단·오류)면 알림을 보낼지'],
  ['REALERT_HOURS', '3', '같은 문제가 계속될 때 몇 시간마다 다시 알릴지'],
  ['EXPIRE_NOTICE_DAYS', '7', '계약 노출 종료 며칠 전부터 만료 임박 알림을 보낼지'],
  ['ALERT_EMAIL', '', '알림 받을 이메일 (쉼표로 여러 개). 비우면 스크립트 속성 또는 본인 계정'],
  ['TOOL_URL', 'https://jguild-app.github.io/TOOLS/', '알림 메일에 넣을 도구 주소'],
  ['MAX_RUN_SECONDS', '300', '한 번 실행에 쓸 최대 시간(초). 넘으면 나머지는 다음 실행에서']
];

var RESULT = { UP: '노출', DOWN: '미노출', OTHER: '타사노출', FAIL: '확인실패' };

var UA = {
  PC: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36',
  M: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1'
};

var NAVER_API_BASE = 'https://api.searchad.naver.com';

/* =========================================================
 * 0. 설치 / 트리거
 * ========================================================= */

/** 최초 1회 실행: 시트와 트리거를 만든다. 여러 번 실행해도 안전. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
    }
  });
  var set = ss.getSheetByName(SHEET.SET);
  var existing = {};
  set.getDataRange().getValues().slice(1).forEach(function (r) { existing[String(r[0])] = true; });
  DEFAULT_SETTINGS.forEach(function (row) { if (!existing[row[0]]) set.appendRow(row); });

  // 기본 시트(Sheet1/시트1)가 비어 있으면 정리
  ['Sheet1', '시트1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  installTriggers_();
  Logger.log('설치 완료. 이제 "광고주" 시트와 "키워드" 시트를 채우거나 05 탭에서 등록하세요.');
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'runCheck' || fn === 'syncContracts') ScriptApp.deleteTrigger(t);
  });
  var interval = Number(getSettings_().CHECK_INTERVAL_MIN || 10);
  if ([1, 5, 10, 15, 30].indexOf(interval) < 0) interval = 10; // Apps Script가 허용하는 값만
  ScriptApp.newTrigger('runCheck').timeBased().everyMinutes(interval).create();
  ScriptApp.newTrigger('syncContracts').timeBased().everyDays(1).atHour(8).create();
}

/* =========================================================
 * 1. 시트 유틸
 * ========================================================= */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트가 없습니다: ' + name + ' — setup()을 먼저 실행하세요.');
  return sh;
}
function rows_(name) {
  var v = sheet_(name).getDataRange().getValues();
  return v.length > 1 ? v.slice(1) : [];
}
function yes_(v) { v = String(v || '').trim().toUpperCase(); return v === '' || v === 'Y' || v === 'YES' || v === 'TRUE' || v === 'O'; }
function nowStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function fmtDate_(d) { return d ? Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''; }
function fmtDateTime_(d) { return d ? Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : ''; }

function getSettings_() {
  var s = {};
  DEFAULT_SETTINGS.forEach(function (r) { s[r[0]] = r[1]; });
  rows_(SHEET.SET).forEach(function (r) { if (r[0]) s[String(r[0]).trim()] = String(r[1]).trim(); });
  var props = PropertiesService.getScriptProperties();
  if (!s.ALERT_EMAIL) s.ALERT_EMAIL = props.getProperty('ALERT_EMAIL') || '';
  if (!s.ALERT_EMAIL) s.ALERT_EMAIL = Session.getEffectiveUser().getEmail();
  return s;
}

function getAdvertisers_() {
  return rows_(SHEET.ADV).filter(function (r) { return String(r[0]).trim(); }).map(function (r, i) {
    return {
      row: i + 2,
      name: String(r[0]).trim(),
      customerId: String(r[1]).replace(/\D/g, ''),
      patterns: String(r[2] || '').split(',').map(function (p) { return p.trim(); }).filter(Boolean),
      email: String(r[3] || '').trim(),
      enabled: yes_(r[4])
    };
  });
}

function getKeywords_() {
  return rows_(SHEET.KW).filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); }).map(function (r, i) {
    return {
      row: i + 2,
      advertiser: String(r[0]).trim(),
      keyword: String(r[1]).trim(),
      media: String(r[2] || 'PC').trim().toUpperCase() === 'M' ? 'M' : 'PC',
      enabled: yes_(r[3]),
      lastStatus: String(r[4] || ''),
      lastChecked: r[5] ? fmtDateTime_(r[5]) : '',
      downStreak: Number(r[6] || 0),
      failStreak: Number(r[7] || 0),
      lastAlert: r[8] ? new Date(r[8]) : null,
      lastEvidence: String(r[9] || '')
    };
  });
}

/* =========================================================
 * 2. 노출 확인 (10분 트리거)
 * ========================================================= */

function runCheck() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('이전 실행이 아직 진행 중이라 건너뜀'); return; }
  try { runCheck_(); } finally { lock.releaseLock(); }
}

function runCheck_() {
  var started = Date.now();
  var settings = getSettings_();
  var maxMs = Number(settings.MAX_RUN_SECONDS || 300) * 1000;
  var advMap = {};
  getAdvertisers_().forEach(function (a) { advMap[a.name] = a; });

  var kws = getKeywords_().filter(function (k) { return k.enabled && advMap[k.advertiser] && advMap[k.advertiser].enabled; });
  if (!kws.length) { Logger.log('확인할 키워드가 없습니다.'); return; }

  // 지난번에 시간이 모자라 못 본 키워드부터 이어서 확인
  var props = PropertiesService.getScriptProperties();
  var cursor = Number(props.getProperty('CHECK_CURSOR') || 0) % kws.length;
  var order = kws.slice(cursor).concat(kws.slice(0, cursor));

  var kwSheet = sheet_(SHEET.KW);
  var logRows = [];
  var alerts = [];
  var done = 0;

  for (var i = 0; i < order.length; i++) {
    if (Date.now() - started > maxMs) { Logger.log('시간 초과로 ' + (order.length - i) + '개는 다음 실행에서 확인'); break; }
    var k = order[i];
    var adv = advMap[k.advertiser];
    var r = checkOne_(k.keyword, k.media, adv.patterns);
    var ts = new Date();

    var downStreak = (r.result === RESULT.DOWN || r.result === RESULT.OTHER) ? k.downStreak + 1 : 0;
    var failStreak = (r.result === RESULT.FAIL) ? k.failStreak + 1 : 0;
    var prevStatus = k.lastStatus;
    var lastAlert = k.lastAlert;

    var alert = decideAlert_(settings, prevStatus, r.result, downStreak, failStreak, lastAlert);
    if (alert) {
      alerts.push({ type: alert, keyword: k, adv: adv, result: r });
      lastAlert = ts;
    }

    kwSheet.getRange(k.row, 5, 1, 6).setValues([[r.result, ts, downStreak, failStreak, lastAlert || '', r.evidence]]);
    logRows.push([ts, k.advertiser, k.keyword, k.media, r.result, r.evidence, r.note || '']);
    done++;
    Utilities.sleep(400); // 네이버에 부담 주지 않도록 간격
  }

  props.setProperty('CHECK_CURSOR', String((cursor + done) % kws.length));
  if (logRows.length) appendLog_(logRows);
  if (alerts.length) sendCheckAlerts_(settings, alerts);
}

/** 한 키워드를 한 매체에서 확인 */
function checkOne_(keyword, media, patterns) {
  var url = media === 'PC'
    ? 'https://search.naver.com/search.naver?where=nexearch&sm=top_hty&query=' + encodeURIComponent(keyword)
    : 'https://m.search.naver.com/search.naver?where=m&sm=mtp_hty&query=' + encodeURIComponent(keyword);
  var res;
  try {
    res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': UA[media],
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8'
      }
    });
  } catch (e) {
    return { result: RESULT.FAIL, evidence: '요청 오류', note: String(e).slice(0, 200) };
  }
  var code = res.getResponseCode();
  var html = res.getContentText('UTF-8') || '';
  if (code !== 200) return { result: RESULT.FAIL, evidence: 'HTTP ' + code, note: html.slice(0, 120).replace(/\s+/g, ' ') };
  if (html.length < 30000 || /captcha|비정상적인\s*접근|자동\s*입력\s*방지|접근이\s*제한/i.test(html.slice(0, 20000))) {
    return { result: RESULT.FAIL, evidence: '차단/비정상 응답 의심 (' + html.length + ' bytes)', note: '' };
  }

  var block = extractBrandBlock_(html, media);
  if (!block) return { result: RESULT.DOWN, evidence: '브랜드검색 영역 없음', note: '' };

  var hit = matchPatterns_(block, patterns);
  if (hit) return { result: RESULT.UP, evidence: '브랜드검색 영역에서 "' + hit + '" 확인', note: '' };

  var head = block.text.slice(0, 80);
  return {
    result: RESULT.OTHER,
    evidence: '브랜드검색 영역은 있으나 식별 문구 불일치 → "' + head + '…"',
    note: patterns.length ? '' : '식별 문구가 비어 있음 — 광고주 시트에 브랜드명을 넣어 주세요'
  };
}

/** 검색 결과 HTML에서 브랜드검색 영역만 잘라 텍스트/링크로 돌려준다 */
function extractBrandBlock_(html, media) {
  var re = media === 'PC'
    ? /<div[^>]+class="[^"]*\bbrand_search\b[^"]*"/
    : /<section[^>]+class="[^"]*\bsp_brand\b[^"]*"/;
  var m = re.exec(html);
  if (!m) return null;
  var seg = html.substr(m.index, 120000);
  var stops = media === 'PC'
    ? ['<div class="sc_new', '<section class="sc_new', 'data-nlog-area="pwl', '<div class="api_subject_bx']
    : ['<section class="sc ', '<div class="api_subject_bx', '<section class="sc_new'];
  var cut = seg.length;
  stops.forEach(function (s) {
    var j = seg.indexOf(s, 1500);
    if (j > 0 && j < cut) cut = j;
  });
  seg = seg.slice(0, cut);
  var text = seg
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;|&amp;|&quot;|&#39;|&lt;|&gt;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  var hrefs = (seg.match(/href="https?:\/\/[^"]+"/g) || []).join(' ');
  var alts = (seg.match(/(?:alt|title|aria-label)="[^"]+"/g) || []).join(' ');
  return { text: text, hrefs: hrefs, alts: alts };
}

function norm_(s) { return String(s || '').toLowerCase().replace(/[\s\-_.,·'"()]/g, ''); }

function matchPatterns_(block, patterns) {
  if (!patterns || !patterns.length) return null;
  var hay = norm_(block.text + ' ' + block.alts) + ' ' + String(block.hrefs).toLowerCase();
  for (var i = 0; i < patterns.length; i++) {
    var p = norm_(patterns[i]);
    if (p && hay.indexOf(p) >= 0) return patterns[i];
  }
  return null;
}

/** 알림을 보낼지 결정: 'DOWN' | 'FAIL' | 'RECOVER' | null */
function decideAlert_(settings, prevStatus, result, downStreak, failStreak, lastAlert) {
  var downAfter = Number(settings.DOWN_ALERT_AFTER || 2);
  var failAfter = Number(settings.FAIL_ALERT_AFTER || 3);
  var realertMs = Number(settings.REALERT_HOURS || 3) * 3600 * 1000;
  var sinceAlert = lastAlert ? Date.now() - lastAlert.getTime() : Infinity;
  var wasProblem = prevStatus === RESULT.DOWN || prevStatus === RESULT.OTHER || prevStatus === RESULT.FAIL;

  if (result === RESULT.UP) {
    // 문제 상태에서 알림까지 나갔던 건이 회복되면 복구 알림
    return (wasProblem && lastAlert) ? 'RECOVER' : null;
  }
  if (result === RESULT.DOWN || result === RESULT.OTHER) {
    if (downStreak === downAfter) return 'DOWN';
    if (downStreak > downAfter && sinceAlert >= realertMs) return 'DOWN';
    return null;
  }
  if (result === RESULT.FAIL) {
    if (failStreak === failAfter) return 'FAIL';
    if (failStreak > failAfter && sinceAlert >= realertMs) return 'FAIL';
    return null;
  }
  return null;
}

function appendLog_(rowsToAdd) {
  var sh = sheet_(SHEET.LOG);
  sh.getRange(sh.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
  var max = 5000; // 로그 시트가 무한정 커지지 않도록
  var extra = sh.getLastRow() - 1 - max;
  if (extra > 0) sh.deleteRows(2, extra);
}

/* =========================================================
 * 3. 이메일 알림
 * ========================================================= */

function recipients_(settings, adv) {
  var list = String(settings.ALERT_EMAIL || '').split(',');
  if (adv && adv.email) list = list.concat(adv.email.split(','));
  var seen = {}, out = [];
  list.forEach(function (e) { e = e.trim(); if (e && !seen[e]) { seen[e] = true; out.push(e); } });
  return out.join(',');
}

function contractSummaryFor_(advName) {
  var list = rows_(SHEET.CONTRACT).filter(function (r) { return String(r[0]).trim() === advName; });
  if (!list.length) return '계약 정보 없음 (계약 동기화를 아직 안 했거나 customerId 미입력)';
  var on = list.filter(function (r) { return r[3] === 'ON_EXPOSING'; });
  if (!on.length) return '⚠ 현재 집행 중(ON_EXPOSING)인 브랜드검색 계약이 없습니다 — 계약 만료 가능성';
  return on.map(function (r) { return r[1] + ' · ' + fmtDate_(r[4]) + ' ~ ' + fmtDate_(r[5]) + ' (D-' + r[6] + ')'; }).join('<br>');
}

function sendCheckAlerts_(settings, alerts) {
  var byAdv = {};
  alerts.forEach(function (a) { (byAdv[a.adv.name] = byAdv[a.adv.name] || []).push(a); });
  Object.keys(byAdv).forEach(function (advName) {
    var items = byAdv[advName];
    var adv = items[0].adv;
    var hasDown = items.some(function (a) { return a.type === 'DOWN'; });
    var hasFail = items.some(function (a) { return a.type === 'FAIL'; });
    var onlyRecover = items.every(function (a) { return a.type === 'RECOVER'; });
    var subject = '[노출 모니터링] ' + (onlyRecover ? '✅ 복구' : hasDown ? '⚠ 미노출' : hasFail ? '❗ 확인 실패' : '알림') + ' · ' + advName +
      ' (' + items.map(function (a) { return a.keyword.keyword + '/' + a.keyword.media; }).join(', ') + ')';

    var lines = items.map(function (a) {
      var label = a.type === 'RECOVER' ? '✅ 다시 노출됨' : a.type === 'FAIL' ? '❗ 확인 실패' : (a.result.result === RESULT.OTHER ? '⚠ 다른 광고주 노출' : '⚠ 미노출');
      return '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;"><b>' + a.keyword.keyword + '</b> <span style="color:#777">(' + a.keyword.media + ')</span></td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;">' + label + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;color:#555;font-size:12px;">' + a.result.evidence + (a.result.note ? '<br>' + a.result.note : '') + '</td></tr>';
    }).join('');

    var html =
      '<div style="font-family:-apple-system,Malgun Gothic,sans-serif;font-size:14px;color:#222;max-width:640px">' +
      '<h2 style="font-size:17px;margin:0 0 12px">' + advName + ' 브랜드검색 노출 알림</h2>' +
      '<p style="color:#555;margin:0 0 14px">' + nowStr_() + ' 확인 결과</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:13px">' + lines + '</table>' +
      '<h3 style="font-size:14px;margin:18px 0 6px">현재 계약 상태</h3>' +
      '<p style="margin:0 0 14px;color:#333">' + contractSummaryFor_(advName) + '</p>' +
      (hasDown ? '<p style="color:#555;font-size:12px;line-height:1.6">미노출 원인은 보통 ① 계약 만료/일시중지 ② 광고그룹·소재 검수 ③ 키워드가 계약에 포함되지 않음 ④ 일시적 네이버 지면 변동입니다. 광고시스템에서 해당 광고그룹 상태를 확인해 주세요.</p>' : '') +
      (hasFail ? '<p style="color:#555;font-size:12px;line-height:1.6">확인 실패는 네이버가 자동 접속을 막았거나 응답이 비정상일 때 기록됩니다. 실제 미노출과는 다르니 직접 검색해 확인해 주세요. 계속되면 확인 간격을 늘리는 것을 권장합니다.</p>' : '') +
      '<p style="margin-top:16px"><a href="' + settings.TOOL_URL + '" style="color:#854D0E">종합 작업 도구 열기 →</a></p></div>';

    var to = recipients_(settings, adv);
    if (to) MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
  });
}

/* =========================================================
 * 4. 네이버 검색광고 API — 브랜드검색 계약 동기화
 * ========================================================= */

function naverApi_(method, path, customerId, query) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('NAVER_API_KEY');
  var secret = props.getProperty('NAVER_API_SECRET');
  if (!key || !secret) throw new Error('스크립트 속성에 NAVER_API_KEY / NAVER_API_SECRET 이 없습니다.');
  var ts = String(Date.now());
  var sigBytes = Utilities.computeHmacSha256Signature(ts + '.' + method + '.' + path, secret, Utilities.Charset.UTF_8);
  var sig = Utilities.base64Encode(sigBytes);
  var qs = '';
  if (query) qs = '?' + Object.keys(query).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(query[k]); }).join('&');
  var res = UrlFetchApp.fetch(NAVER_API_BASE + path + qs, {
    method: method.toLowerCase(),
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'X-Timestamp': ts,
      'X-API-KEY': key,
      'X-Customer': String(customerId),
      'X-Signature': sig
    }
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code >= 300) throw new Error('네이버 API ' + code + ' (' + path + ', customer ' + customerId + '): ' + body.slice(0, 200));
  return body ? JSON.parse(body) : null;
}

/** 모든 광고주의 브랜드검색 계약을 가져와 "계약" 시트를 새로 쓴다 (하루 1회 트리거 + 수동) */
function syncContracts() {
  var settings = getSettings_();
  var advs = getAdvertisers_().filter(function (a) { return a.enabled && a.customerId; });
  var out = [];
  var errors = [];
  var now = Date.now();
  var expiring = [];

  advs.forEach(function (adv) {
    var list;
    try { list = naverApi_('GET', '/ncc/time-contracts', adv.customerId) || []; }
    catch (e) { errors.push(adv.name + ': ' + String(e.message || e)); return; }

    // 집행 중 / 예정 / 취소 예정 + 최근 30일 내 종료된 계약만 보여준다
    var keep = list.filter(function (c) {
      var st = c.contractStatus;
      if (st === 'ON_EXPOSING' || st === 'UPCOMING_EXPOSE' || st === 'UPCOMING_CANCEL') return true;
      if (st === 'EXPOSE_COMPLETED' && c.exposureEndDt && (now - new Date(c.exposureEndDt).getTime()) < 30 * 86400000) return true;
      return false;
    });
    keep.sort(function (a, b) { return String(a.adgroupName).localeCompare(String(b.adgroupName)) || String(a.exposureStartDt).localeCompare(String(b.exposureStartDt)); });

    var upcomingByGroup = {};
    keep.forEach(function (c) { if (c.contractStatus === 'UPCOMING_EXPOSE') upcomingByGroup[c.nccAdgroupId] = true; });

    keep.forEach(function (c) {
      var end = c.exposureEndDt ? new Date(c.exposureEndDt) : null;
      var daysLeft = end ? Math.ceil((end.getTime() - now) / 86400000) : '';
      out.push([adv.name, c.adgroupName || '', c.contractName || '', c.contractStatus || '', c.exposureStartDt ? new Date(c.exposureStartDt) : '', end || '', daysLeft, c.nccTimeContractId || '', new Date()]);
      if (c.contractStatus === 'ON_EXPOSING' && daysLeft !== '' && daysLeft <= Number(settings.EXPIRE_NOTICE_DAYS || 7) && !upcomingByGroup[c.nccAdgroupId]) {
        expiring.push({ adv: adv, c: c, daysLeft: daysLeft });
      }
    });
    if (!keep.length) out.push([adv.name, '', '(조회된 브랜드검색 계약 없음)', '', '', '', '', '', new Date()]);
  });

  var sh = sheet_(SHEET.CONTRACT);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, out[0].length).setValues(out);
  if (errors.length) sh.getRange(2 + out.length, 1, 1, 3).setValues([['(오류)', errors.join(' / '), '']]);

  if (expiring.length) sendExpiryAlerts_(settings, expiring);
  if (errors.length) Logger.log('계약 동기화 오류: ' + errors.join(' / '));
  return { count: out.length, errors: errors };
}

function sendExpiryAlerts_(settings, expiring) {
  var props = PropertiesService.getScriptProperties();
  var today = fmtDate_(new Date());
  var byAdv = {};
  expiring.forEach(function (x) {
    var key = 'EXPIRE_SENT_' + x.c.nccTimeContractId;
    if (props.getProperty(key) === today) return; // 하루 한 번만
    props.setProperty(key, today);
    (byAdv[x.adv.name] = byAdv[x.adv.name] || []).push(x);
  });
  Object.keys(byAdv).forEach(function (name) {
    var items = byAdv[name];
    var subject = '[노출 모니터링] ⏰ 계약 만료 임박 · ' + name + ' (' + items.map(function (x) { return 'D-' + x.daysLeft; }).join(', ') + ')';
    var html = '<div style="font-family:-apple-system,Malgun Gothic,sans-serif;font-size:14px;color:#222;max-width:640px">' +
      '<h2 style="font-size:17px;margin:0 0 12px">' + name + ' 브랜드검색 계약이 곧 끝납니다</h2>' +
      '<ul style="line-height:1.8">' + items.map(function (x) {
        return '<li><b>' + x.c.adgroupName + '</b> · ' + fmtDate_(x.c.exposureStartDt) + ' ~ <b>' + fmtDate_(x.c.exposureEndDt) + '</b> (D-' + x.daysLeft + ')<br><span style="color:#777;font-size:12px">같은 광고그룹에 예정(UPCOMING_EXPOSE) 계약이 없습니다.</span></li>';
      }).join('') + '</ul>' +
      '<p><a href="' + settings.TOOL_URL + '" style="color:#854D0E">종합 작업 도구 열기 →</a></p></div>';
    var to = recipients_(settings, items[0].adv);
    if (to) MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
  });
}

/* =========================================================
 * 5. 웹앱 (index.html 05 탭과 통신)
 * ========================================================= */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty('WEB_TOKEN');
  if (!expected) throw new Error('스크립트 속성에 WEB_TOKEN 이 설정되어 있지 않습니다.');
  if (String(token || '') !== expected) throw new Error('접속 토큰이 올바르지 않습니다.');
}

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    checkToken_(p.token);
    return json_(statusPayload_());
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    checkToken_(body.token);
    var r = handleAction_(body);
    var payload = statusPayload_();
    payload.actionResult = r || null;
    return json_(payload);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function statusPayload_() {
  var advs = getAdvertisers_();
  var kws = getKeywords_();
  var contracts = rows_(SHEET.CONTRACT).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    return { advertiser: r[0], adgroupName: r[1], contractName: r[2], status: r[3], exposureStartDt: fmtDate_(r[4]), exposureEndDt: fmtDate_(r[5]), daysLeft: r[6] === '' ? null : Number(r[6]), contractId: r[7], updatedAt: fmtDateTime_(r[8]) };
  });
  var logAll = rows_(SHEET.LOG);
  var log = logAll.slice(Math.max(0, logAll.length - 80)).reverse().map(function (r) {
    return { time: fmtDateTime_(r[0]), advertiser: r[1], keyword: r[2], media: r[3], result: r[4], evidence: r[5], note: r[6] };
  });
  var triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  return {
    ok: true,
    serverTime: nowStr_(),
    triggersInstalled: triggers.indexOf('runCheck') >= 0,
    settings: getSettings_(),
    advertisers: advs.map(function (a) { return { name: a.name, customerId: a.customerId, patterns: a.patterns, email: a.email, enabled: a.enabled }; }),
    keywords: kws.map(function (k) { return { advertiser: k.advertiser, keyword: k.keyword, media: k.media, enabled: k.enabled, lastStatus: k.lastStatus, lastChecked: k.lastChecked, downStreak: k.downStreak, failStreak: k.failStreak, lastEvidence: k.lastEvidence }; }),
    contracts: contracts,
    log: log
  };
}

function handleAction_(b) {
  var action = b.action;
  if (action === 'addAdvertiser') {
    if (!b.name) throw new Error('광고주명이 비어 있습니다.');
    var exists = getAdvertisers_().some(function (a) { return a.name === String(b.name).trim(); });
    if (exists) throw new Error('이미 있는 광고주명입니다: ' + b.name);
    sheet_(SHEET.ADV).appendRow([String(b.name).trim(), String(b.customerId || '').replace(/\D/g, ''), String(b.patterns || b.name).trim(), String(b.email || '').trim(), 'Y']);
    return { message: '광고주 추가됨: ' + b.name };
  }
  if (action === 'updateAdvertiser') {
    var adv = getAdvertisers_().filter(function (a) { return a.name === b.name; })[0];
    if (!adv) throw new Error('광고주를 찾을 수 없습니다: ' + b.name);
    sheet_(SHEET.ADV).getRange(adv.row, 2, 1, 4).setValues([[String(b.customerId || adv.customerId).replace(/\D/g, ''), b.patterns != null ? String(b.patterns) : adv.patterns.join(', '), b.email != null ? String(b.email) : adv.email, b.enabled === false ? 'N' : 'Y']]);
    return { message: '광고주 수정됨: ' + b.name };
  }
  if (action === 'removeAdvertiser') {
    var a2 = getAdvertisers_().filter(function (a) { return a.name === b.name; })[0];
    if (!a2) throw new Error('광고주를 찾을 수 없습니다: ' + b.name);
    // 키워드도 같이 삭제 (아래에서 위로)
    var ks = getKeywords_().filter(function (k) { return k.advertiser === b.name; }).sort(function (x, y) { return y.row - x.row; });
    ks.forEach(function (k) { sheet_(SHEET.KW).deleteRow(k.row); });
    sheet_(SHEET.ADV).deleteRow(a2.row);
    return { message: '광고주와 키워드 ' + ks.length + '개 삭제됨: ' + b.name };
  }
  if (action === 'addKeyword') {
    if (!b.advertiser || !b.keyword) throw new Error('광고주와 키워드를 모두 입력하세요.');
    var media = String(b.media || 'PC').toUpperCase();
    var medias = media === 'ALL' ? ['PC', 'M'] : [media === 'M' ? 'M' : 'PC'];
    var existing = getKeywords_();
    var added = 0;
    var kwList = String(b.keyword).split(/[\n,]/).map(function (s) { return s.trim(); }).filter(Boolean);
    kwList.forEach(function (kw) {
      medias.forEach(function (md) {
        var dup = existing.some(function (k) { return k.advertiser === b.advertiser && k.keyword === kw && k.media === md; });
        if (dup) return;
        sheet_(SHEET.KW).appendRow([b.advertiser, kw, md, 'Y', '', '', 0, 0, '', '']);
        added++;
      });
    });
    return { message: '키워드 ' + added + '개 추가됨' };
  }
  if (action === 'removeKeyword') {
    var k2 = getKeywords_().filter(function (k) { return k.advertiser === b.advertiser && k.keyword === b.keyword && k.media === b.media; })[0];
    if (!k2) throw new Error('키워드를 찾을 수 없습니다.');
    sheet_(SHEET.KW).deleteRow(k2.row);
    return { message: '키워드 삭제됨: ' + b.keyword + ' (' + b.media + ')' };
  }
  if (action === 'toggleKeyword') {
    var k3 = getKeywords_().filter(function (k) { return k.advertiser === b.advertiser && k.keyword === b.keyword && k.media === b.media; })[0];
    if (!k3) throw new Error('키워드를 찾을 수 없습니다.');
    sheet_(SHEET.KW).getRange(k3.row, 4).setValue(k3.enabled ? 'N' : 'Y');
    return { message: (k3.enabled ? '일시중지' : '재개') + ': ' + b.keyword };
  }
  if (action === 'runCheck') {
    runCheck();
    return { message: '노출 확인 실행 완료' };
  }
  if (action === 'syncContracts') {
    var r = syncContracts();
    return { message: '계약 동기화 완료 (' + r.count + '건' + (r.errors.length ? ', 오류 ' + r.errors.length + '건: ' + r.errors.join(' / ') : '') + ')' };
  }
  if (action === 'testEmail') {
    var s = getSettings_();
    MailApp.sendEmail({ to: recipients_(s, null), subject: '[노출 모니터링] 테스트 메일', htmlBody: '알림 메일이 정상적으로 도착했습니다. (' + nowStr_() + ')' });
    return { message: '테스트 메일 발송: ' + recipients_(s, null) };
  }
  if (action === 'setSetting') {
    var set = sheet_(SHEET.SET);
    var vals = set.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(b.key)) { set.getRange(i + 1, 2).setValue(String(b.value)); return { message: '설정 저장: ' + b.key }; }
    }
    set.appendRow([String(b.key), String(b.value), '']);
    return { message: '설정 추가: ' + b.key };
  }
  if (action === 'reinstallTriggers') {
    installTriggers_();
    return { message: '확인 간격 ' + getSettings_().CHECK_INTERVAL_MIN + '분으로 트리거 재설치 완료' };
  }
  if (action === 'ping') return { message: 'ok' };
  throw new Error('알 수 없는 action: ' + action);
}

/* =========================================================
 * 6. 수동 테스트용
 * ========================================================= */

/** 편집기에서 바로 실행해 특정 키워드 판정을 눈으로 확인 */
function debugCheck() {
  var r1 = checkOne_('올리브영', 'PC', ['올리브영']);
  var r2 = checkOne_('올리브영', 'M', ['올리브영']);
  Logger.log(JSON.stringify(r1));
  Logger.log(JSON.stringify(r2));
}

/** API 연결 확인: 첫 번째 광고주의 계약 목록을 로그로 출력 */
function debugApi() {
  var adv = getAdvertisers_().filter(function (a) { return a.customerId; })[0];
  if (!adv) { Logger.log('customerId가 있는 광고주가 없습니다.'); return; }
  var list = naverApi_('GET', '/ncc/time-contracts', adv.customerId);
  Logger.log(adv.name + ': ' + (list ? list.length : 0) + '건');
  (list || []).slice(0, 10).forEach(function (c) { Logger.log([c.adgroupName, c.contractStatus, c.exposureStartDt, c.exposureEndDt].join(' | ')); });
}
