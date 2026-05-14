/**
 * 청개구리 사전점검 GAS Backend v1.2
 * ================================================================
 * 사전점검 앱 (https://frogsash.co.kr/inspect/ + SPARE + frogcheck) 공통 백엔드.
 *
 * v1.2 보안 강화 (2026-05-15):
 *  - handleAi에 mode 파라미터 추가 ('diag' 진단 / 'chat' 챗봇)
 *  - mode='diag'에 서버 측 quota 강제 (FREE 2회 / PAID 100회) — 클라이언트 우회 불가
 *  - prompts 배열 길이 제한 (diag 2, chat 1)
 *  - 짧은 수명 nonce 토큰 흐름 (30초 TTL, 1회용, fp 바인딩, HMAC 서명)
 *  - AiLog 시트에 호출 기록 + 의심 호출 추적
 *  - getUserPaidMode_(fp): 오늘 유료 코드 활성화 여부로 free/paid 판별
 *
 * ── 1회 셋업 (5분) ────────────────────────────────────────────
 *  1. 새 Apps Script 프로젝트 생성: 이름 "청개구리-사전점검-Backend"
 *  2. 이 파일 전체를 Code.gs에 붙여넣기
 *  3. 좌측 ⚙️ "프로젝트 설정" → 하단 "스크립트 속성" → 다음 4개 추가:
 *       - SHEET_ID         : 코드/사용량 시트 문서 ID (setup() 실행 시 자동 생성)
 *       - GEMINI_API_KEY   : Google AI Studio 발급 진짜 Gemini 키
 *       - ADMIN_PASSWORD   : 관리자 모드 비밀번호 (자유 설정, 8자 이상 권장)
 *       - HMAC_SECRET      : 토큰 서명용 임의 문자열 (32자 이상 무작위)
 *  4. setup() 한 번 실행 → 시트 자동 생성 + ID 자동 저장
 *  5. "배포" → "새 배포" → 유형: 웹 앱 (모든 사용자 / 나로 실행)
 *  6. 웹 앱 URL 복사 → index.html의 GAS_URL 상수 교체 후 git push
 *
 * ── 코드 발급 ─────────────────────────────────────────────────
 *  방법 A) addCode() 함수 직접 실행 (ADD_CODE_INPUT 채우기)
 *  방법 B) 시트 "Codes"에 직접 한 줄 추가 (code 열에 대문자 4~10자)
 *  방법 C) bulkAddCodes(20) → 미사용 코드 20개 풀
 *
 * ── 보안 요약 ─────────────────────────────────────────────────
 *  - Gemini 키 서버 보관 (클라이언트 노출 0)
 *  - 관리자 토큰: HMAC 서명 + 24h TTL (stateless)
 *  - nonce 토큰: HMAC 서명 + 30s TTL + 1회용 (fp 바인딩) — AI 무한 호출 차단
 *  - 인증코드: 1회용 + 기기 fp 바인딩 + 당일 자정 만료
 *  - AI 진단 quota 서버 강제: FREE 2회/일, PAID 100회/일
 * ================================================================
 */

// ───── 상수 ─────
const FREE_LIMIT = 2;
const PAID_LIMIT = 100;
const CHAT_LIMIT = 30;
const ADMIN_LIMIT = 99999;
const ADMIN_TOKEN_TTL_MS = 24 * 60 * 60 * 1000;
const NONCE_TTL_MS = 30 * 1000; // 30초
const DIAG_MAX_PROMPTS = 2;
const CHAT_MAX_PROMPTS = 1;
const VERSION = 'v1.2';
const GEMINI_MODEL = 'gemini-3.1-flash-lite';

// addCode() 임시 입력 (편집기에서 직접 실행 시 사용)
const ADD_CODE_INPUT = {
  code: '',           // 비워두면 자동 생성
  amount: 10000,
  phone: '',
  note: ''
};

// ───── 진입점 ─────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'verify') return jsonRes(handleVerify(e.parameter));
    if (action === 'usage')  return jsonRes(handleUsage(e.parameter));
    return jsonRes({ status: 'ok', service: '청개구리-사전점검-Backend', version: VERSION });
  } catch (err) {
    return jsonRes({ valid: false, error: 'server_error', message: '서버 오류: ' + err.message });
  }
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents || '{}');
  } catch (err) {
    return jsonRes({ error: 'bad_json', message: '잘못된 요청 형식' });
  }
  const action = body.action || '';
  try {
    if (action === 'ping')        return jsonRes({ ok: true, version: VERSION });
    if (action === 'nonce')       return jsonRes(handleNonce(body));
    if (action === 'use')         return jsonRes(handleUse(body));
    if (action === 'verify')      return jsonRes(handleVerify(body));
    if (action === 'usage')       return jsonRes(handleUsage(body));
    if (action === 'admin_auth')  return jsonRes(handleAdminAuth(body));
    if (action === 'ai')          return jsonRes(handleAi(body));
    return jsonRes({ error: 'unknown_action', valid_actions: ['ping','nonce','use','verify','usage','admin_auth','ai'] });
  } catch (err) {
    return jsonRes({ error: 'server_error', message: '서버 오류: ' + err.message });
  }
}

// ───── 핵심 액션 ─────

// 코드 검증: 1회용, 기기 바인딩, 당일 자정 만료
function handleVerify(p) {
  const code = String(p.code || '').trim().toUpperCase();
  const fp = String(p.fp || '').trim();
  if (code.length < 4) return { valid: false, message: '코드를 입력해주세요 (4자 이상)' };
  if (!fp) return { valid: false, message: '기기 정보가 없습니다' };

  const sh = sheet_('Codes');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { valid: false, message: '존재하지 않는 코드입니다' };

  const data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[0]).toUpperCase() !== code) continue;
    const rowNum = i + 2;
    const status = String(r[6] || 'active').toLowerCase();

    if (status === 'refunded')
      return { valid: false, message: '환불 처리된 코드입니다' };

    if (status === 'used') {
      const usedFp = String(r[5] || '');
      const usedAt = r[4] ? new Date(r[4]) : null;
      const today = ymd_(new Date());
      const usedDay = usedAt ? ymd_(usedAt) : '';

      if (usedDay !== today)
        return { valid: false, message: '만료된 코드입니다 (당일만 사용 가능)' };
      if (usedFp && usedFp !== fp)
        return { valid: false, message: '다른 기기에서 사용된 코드입니다' };
      return { valid: true, message: '재인증 성공' };
    }

    // active → 사용 처리
    const now = new Date();
    sh.getRange(rowNum, 5).setValue(now);
    sh.getRange(rowNum, 6).setValue(fp);
    sh.getRange(rowNum, 7).setValue('used');
    SpreadsheetApp.flush();
    return { valid: true, message: '인증 성공' };
  }
  return { valid: false, message: '존재하지 않는 코드입니다' };
}

// nonce 발급: 30초 TTL + fp 바인딩 + HMAC 서명 + 1회용 (Cache 추적)
function handleNonce(p) {
  const fp = String(p.fp || '').trim();
  if (!fp) return { error: 'no_fp', message: '기기 정보 없음' };
  return { nonce: makeNonce_(fp) };
}

function handleUsage(p) {
  const fp = String(p.fp || '').trim();
  const mode = String(p.mode || 'free');
  if (!fp) return { count: 0 };
  const today = ymd_(new Date());
  const row = findUsageRow_(fp, mode, today);
  return { count: row ? Number(row.count || 0) : 0 };
}

function handleUse(p) {
  const fp = String(p.fp || '').trim();
  const mode = String(p.mode || 'free');
  if (!fp) return { ok: false };
  const today = ymd_(new Date());
  upsertUsage_(fp, mode, today, 1);
  pruneOldUsage_();
  return { ok: true };
}

function handleAdminAuth(p) {
  const pwd = String(p.pwd || '');
  const fp = String(p.fp || '');
  const expected = props_().getProperty('ADMIN_PASSWORD') || '';
  if (!expected) return { success: false, message: 'ADMIN_PASSWORD 미설정' };
  if (pwd !== expected) return { success: false };
  const token = makeAdminToken_(fp);
  return { success: true, token: token };
}

// AI 호출 (Gemini 프록시) — v1.2 강화
function handleAi(p) {
  const fp = String(p.fp || '').trim();
  const mode = String(p.mode || 'chat').toLowerCase(); // 'diag' or 'chat'
  const nonce = String(p.nonce || '');
  const adminToken = String(p.adminToken || '');
  const prompts = Array.isArray(p.prompts) ? p.prompts : [];

  if (!fp) return { error: 'no_fp', message: '기기 정보 없음' };
  if (!prompts.length) return { error: 'no_prompts', message: '프롬프트 없음' };
  if (mode !== 'diag' && mode !== 'chat')
    return { error: 'bad_mode', message: '잘못된 mode (diag 또는 chat)' };

  const isAdmin = adminToken && verifyAdminToken_(adminToken);

  // nonce 검증 (관리자 면제) — AI 무한 호출 차단의 핵심
  if (!isAdmin && !verifyNonce_(nonce, fp)) {
    logAi_(fp, mode, prompts[0] && prompts[0].prompt, 0, 'invalid_nonce');
    return { error: 'invalid_nonce', message: '인증이 만료됐습니다. 페이지를 새로고침해주세요' };
  }

  // 프롬프트 개수 제한 (DOS 방어)
  const maxPrompts = (mode === 'diag') ? DIAG_MAX_PROMPTS : CHAT_MAX_PROMPTS;
  if (prompts.length > maxPrompts) {
    logAi_(fp, mode, prompts[0] && prompts[0].prompt, 0, 'too_many_prompts');
    return { error: 'too_many_prompts', message: '한 번에 최대 ' + maxPrompts + '개까지 호출 가능' };
  }

  // quota 검증 (관리자 면제)
  const today = ymd_(new Date());
  let userQuotaMode = '';
  if (!isAdmin) {
    if (mode === 'diag') {
      userQuotaMode = getUserPaidMode_(fp); // 'paid' or 'free'
      const u = findUsageRow_(fp, userQuotaMode, today);
      const cur = u ? Number(u.count || 0) : 0;
      const limit = (userQuotaMode === 'paid') ? PAID_LIMIT : FREE_LIMIT;
      if (cur >= limit) {
        logAi_(fp, mode, prompts[0] && prompts[0].prompt, 0, 'diag_limit');
        return { error: 'diag_limit', message: '진단 일일 한도 도달 (' + cur + '/' + limit + ')', userMode: userQuotaMode };
      }
    } else {
      const u = findUsageRow_(fp, 'chat', today);
      const cur = u ? Number(u.count || 0) : 0;
      if (cur >= CHAT_LIMIT) {
        logAi_(fp, mode, prompts[0] && prompts[0].prompt, 0, 'chat_limit');
        return { error: 'chat_limit', message: '챗봇 일일 한도 도달 (' + CHAT_LIMIT + '회)' };
      }
    }
  }

  const apiKey = props_().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error: 'no_api_key', message: 'GEMINI_API_KEY 미설정' };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  const results = [];
  let totalChars = 0;
  for (let i = 0; i < prompts.length; i++) {
    const item = prompts[i] || {};
    const parts = [{ text: String(item.prompt || '') }];
    const images = Array.isArray(item.images) ? item.images : [];
    images.forEach(img => {
      if (img && img.data && img.mimeType)
        parts.push({ inline_data: { mime_type: img.mimeType, data: img.data } });
    });
    const defaultConfig = { temperature: 0.3, maxOutputTokens: 1024 };
    const payload = {
      contents: [{ parts: parts }],
      generationConfig: Object.assign(defaultConfig, item.config || {})
    };
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code !== 200) {
        results.push('AI 응답 오류 (' + code + ')');
        continue;
      }
      const data = JSON.parse(res.getContentText('UTF-8'));
      const txt = (data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0] && data.candidates[0].content.parts[0].text) || '';
      results.push(txt);
      totalChars += txt.length;
    } catch (err) {
      results.push('AI 호출 실패: ' + err.message);
    }
  }

  // 성공 시 quota +1 + 로깅
  if (!isAdmin) {
    if (mode === 'diag') {
      upsertUsage_(fp, userQuotaMode, today, 1);
    } else {
      upsertUsage_(fp, 'chat', today, 1);
    }
    logAi_(fp, mode, prompts[0] && prompts[0].prompt, totalChars, 'ok');
  }
  return { results: results };
}

// 코드/사용자 모드 판별: 오늘 paid 코드 활성화 여부
function getUserPaidMode_(fp) {
  const sh = sheet_('Codes');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'free';
  const today = ymd_(new Date());
  const data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const status = String(r[6] || '').toLowerCase();
    if (status !== 'used') continue;
    const usedFp = String(r[5] || '');
    const usedAt = r[4] ? new Date(r[4]) : null;
    const usedDay = usedAt ? ymd_(usedAt) : '';
    if (usedFp === fp && usedDay === today) return 'paid';
  }
  return 'free';
}

// ───── 시트 헬퍼 ─────

function spreadsheet_() {
  let id = props_().getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('청개구리-사전점검-Backend-Data');
    id = ss.getId();
    props_().setProperty('SHEET_ID', id);
  }
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === 'Codes') {
      sh.getRange(1, 1, 1, 8).setValues([['code','issued_at','amount','phone','used_at','used_fp','status','note']]);
      sh.getRange(1, 1, 1, 8).setFontWeight('bold');
      sh.setFrozenRows(1);
    } else if (name === 'Usage') {
      sh.getRange(1, 1, 1, 5).setValues([['fp','mode','date','count','last_seen']]);
      sh.getRange(1, 1, 1, 5).setFontWeight('bold');
      sh.setFrozenRows(1);
    } else if (name === 'AiLog') {
      sh.getRange(1, 1, 1, 5).setValues([['ts','fp','prompt_first','result_chars','status']]);
      sh.getRange(1, 1, 1, 5).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function findUsageRow_(fp, mode, date) {
  const sh = sheet_('Usage');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === fp && String(data[i][1]) === mode && ymd_(data[i][2]) === date) {
      return { rowNum: i + 2, fp: data[i][0], mode: data[i][1], date: data[i][2], count: data[i][3], last_seen: data[i][4] };
    }
  }
  return null;
}

function upsertUsage_(fp, mode, date, delta) {
  const sh = sheet_('Usage');
  const existing = findUsageRow_(fp, mode, date);
  const now = new Date();
  if (existing) {
    sh.getRange(existing.rowNum, 4).setValue(Number(existing.count || 0) + delta);
    sh.getRange(existing.rowNum, 5).setValue(now);
  } else {
    sh.appendRow([fp, mode, date, delta, now]);
  }
  SpreadsheetApp.flush();
}

function pruneOldUsage_() {
  const sh = sheet_('Usage');
  const lastRow = sh.getLastRow();
  const MAX = 1000;
  if (lastRow - 1 <= MAX) return;
  const excess = (lastRow - 1) - MAX;
  sh.deleteRows(2, excess);
}

// AI 호출 로깅 (1000건 자동 prune)
function logAi_(fp, mode, promptText, resultChars, status) {
  try {
    const sh = sheet_('AiLog');
    sh.appendRow([new Date(), fp, String(promptText || '').slice(0, 80), resultChars, mode + '|' + status]);
    const lastRow = sh.getLastRow();
    if (lastRow - 1 > 1000) sh.deleteRows(2, lastRow - 1 - 1000);
  } catch (e) {}
}

// ───── 토큰 (HMAC) ─────

function makeAdminToken_(fp) {
  const exp = Date.now() + ADMIN_TOKEN_TTL_MS;
  const payload = fp + '|' + exp;
  const sig = hmac_(payload);
  return Utilities.base64EncodeWebSafe(payload) + '.' + sig;
}

function verifyAdminToken_(token) {
  try {
    const parts = String(token).split('.');
    if (parts.length !== 2) return false;
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    const sigCheck = hmac_(payload);
    if (sigCheck !== parts[1]) return false;
    const exp = Number(payload.split('|')[1] || 0);
    return Date.now() < exp;
  } catch (e) { return false; }
}

// nonce: fp 바인딩 + 30초 TTL + HMAC 서명 + 1회용 (Cache로 중복 차단)
function makeNonce_(fp) {
  const exp = Date.now() + NONCE_TTL_MS;
  const r = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const payload = fp + '|' + exp + '|' + r;
  const sig = hmac_(payload);
  return Utilities.base64EncodeWebSafe(payload) + '.' + sig;
}

function verifyNonce_(token, fp) {
  try {
    if (!token) return false;
    const parts = String(token).split('.');
    if (parts.length !== 2) return false;
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    const sigCheck = hmac_(payload);
    if (sigCheck !== parts[1]) return false;
    const segs = payload.split('|');
    if (segs[0] !== fp) return false;
    const exp = Number(segs[1] || 0);
    if (Date.now() > exp) return false;
    // 1회용: Cache로 중복 사용 차단 (60초간 nonce 추적)
    const cache = CacheService.getScriptCache();
    const key = 'nonce_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, payload).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
    if (cache.get(key)) return false;
    cache.put(key, '1', 60);
    return true;
  } catch (e) { return false; }
}

function hmac_(text) {
  const secret = props_().getProperty('HMAC_SECRET') || 'CHANGE_ME';
  const sig = Utilities.computeHmacSha256Signature(text, secret);
  return Utilities.base64EncodeWebSafe(sig);
}

// ───── 유틸 ─────

function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ymd_(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  return Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
}

function props_() { return PropertiesService.getScriptProperties(); }

// ───── 셋업 / 운영 도구 ─────

function setup() {
  sheet_('Codes');
  sheet_('Usage');
  sheet_('AiLog');
  const id = props_().getProperty('SHEET_ID');
  Logger.log('✅ Setup OK. SHEET_ID=' + id);
  Logger.log('스크립트 속성: GEMINI_API_KEY, ADMIN_PASSWORD, HMAC_SECRET 설정 확인');
  return id;
}

function addCode() {
  const code = (ADD_CODE_INPUT.code || randCode_(6)).toUpperCase();
  const sh = sheet_('Codes');
  sh.appendRow([code, new Date(), ADD_CODE_INPUT.amount || 10000, ADD_CODE_INPUT.phone || '', '', '', 'active', ADD_CODE_INPUT.note || '']);
  Logger.log('✅ 발급된 코드: ' + code);
  return code;
}

function bulkAddCodes(n) {
  n = n || 10;
  const sh = sheet_('Codes');
  const codes = [];
  for (let i = 0; i < n; i++) {
    const c = randCode_(6);
    sh.appendRow([c, new Date(), 10000, '', '', '', 'active', 'pool']);
    codes.push(c);
  }
  Logger.log('✅ 발급된 코드 ' + n + '개:\n' + codes.join('\n'));
  return codes;
}

const REFUND_CODE = '';
function refundCode() {
  const code = REFUND_CODE.toUpperCase();
  if (!code) { Logger.log('REFUND_CODE 상수에 코드를 먼저 입력하세요'); return; }
  const sh = sheet_('Codes');
  const lastRow = sh.getLastRow();
  const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === code) {
      sh.getRange(i + 2, 7).setValue('refunded');
      Logger.log('✅ 환불 처리: ' + code);
      return;
    }
  }
  Logger.log('❌ 코드 없음: ' + code);
}

function randCode_(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let out = '';
  for (let i = 0; i < len; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}
