/**
 * 청개구리 사전점검 GAS Backend v1.0
 * ================================================================
 * 사전점검 앱 (https://frogsash.co.kr/inspect/) 전용 백엔드.
 * 클라이언트(index.html) 변경 없이 GAS_URL만 새 배포 URL로 교체하면 작동.
 *
 * ── 1회 셋업 (5분) ────────────────────────────────────────────
 *  1. 새 Apps Script 프로젝트 생성: 이름 "청개구리-사전점검-Backend"
 *  2. 이 파일 전체를 Code.gs에 붙여넣기
 *  3. 좌측 ⚙️ "프로젝트 설정" → 하단 "스크립트 속성" → 다음 4개 추가:
 *       - SHEET_ID         : 코드/사용량 시트 문서 ID (아래 setup() 실행 후 자동 생성도 가능)
 *       - GEMINI_API_KEY   : Google AI Studio 발급 진짜 Gemini 키
 *       - ADMIN_PASSWORD   : 관리자 모드 비밀번호 (자유 설정, 8자 이상 권장)
 *       - HMAC_SECRET      : 토큰 서명용 임의 문자열 (32자 이상 무작위)
 *  4. SHEET_ID를 비워두고 setup()을 한 번 실행 → 시트 자동 생성 + ID 자동 저장
 *  5. ▶ 실행 메뉴에서 setup() 한 번 실행해 권한 승인 (Spreadsheet, UrlFetch)
 *  6. 우상단 "배포" → "새 배포" → 유형: 웹 앱
 *       - 액세스: 모든 사용자
 *       - 다음 사용자 자격으로 실행: 나
 *     → 배포 → 웹 앱 URL 복사
 *  7. D:/github/inspect/index.html의 GAS_URL 상수를 새 URL로 교체 후 git push
 *
 * ── 코드 발급 (사장님이 결제 받을 때마다) ──────────────────
 *  방법 A) GAS 편집기 → 함수 드롭다운 → addCode → 실행
 *           코드/금액/연락처/메모를 ADD_CODE_INPUT 상수에 임시 입력
 *  방법 B) 시트 "Codes"에 직접 한 줄 추가 (code 열에 대문자 4~10자)
 *  방법 C) bulkAddCodes(20) 실행 → 미사용 코드 20개 미리 풀로 만들어두기
 *
 *  발급된 코드를 카카오톡/문자로 구매자에게 전달하면 끝.
 *
 * ── 보안 ──────────────────────────────────────────────────────
 *  - Gemini 키는 서버에만 저장, 클라이언트에 노출 안됨
 *  - 관리자 토큰은 HMAC 서명 + 24시간 만료 (stateless)
 *  - 코드는 1회용 + 기기 fingerprint 바인딩 (다른 폰에서 재사용 불가)
 * ================================================================
 */

// ───── 상수 ─────
const FREE_LIMIT = 2;
const PAID_LIMIT = 100;
const ADMIN_LIMIT = 99999;
const ADMIN_TOKEN_TTL_MS = 24 * 60 * 60 * 1000;
const VERSION = 'v1.0';
const GEMINI_MODEL = 'gemini-2.5-flash-lite';

// addCode() 임시 입력 (편집기에서 직접 실행 시 사용)
const ADD_CODE_INPUT = {
  code: '',           // 예: 'A7K3X9' (비워두면 자동 생성)
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
    if (action === 'use')         return jsonRes(handleUse(body));
    if (action === 'verify')      return jsonRes(handleVerify(body));
    if (action === 'usage')       return jsonRes(handleUsage(body));
    if (action === 'admin_auth')  return jsonRes(handleAdminAuth(body));
    if (action === 'ai')          return jsonRes(handleAi(body));
    return jsonRes({ error: 'unknown_action', valid_actions: ['ping','use','verify','usage','admin_auth','ai'] });
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
      // 같은 기기 + 같은 날 → 멱등 재인증
      return { valid: true, message: '재인증 성공' };
    }

    // active → 사용 처리
    const now = new Date();
    sh.getRange(rowNum, 5).setValue(now);     // E: used_at
    sh.getRange(rowNum, 6).setValue(fp);      // F: used_fp
    sh.getRange(rowNum, 7).setValue('used');  // G: status
    SpreadsheetApp.flush();
    return { valid: true, message: '인증 성공' };
  }
  return { valid: false, message: '존재하지 않는 코드입니다' };
}

// 사용량 조회
function handleUsage(p) {
  const fp = String(p.fp || '').trim();
  const mode = String(p.mode || 'free');
  if (!fp) return { count: 0 };
  const today = ymd_(new Date());
  const row = findUsageRow_(fp, mode, today);
  return { count: row ? Number(row.count || 0) : 0 };
}

// 사용량 +1
function handleUse(p) {
  const fp = String(p.fp || '').trim();
  const mode = String(p.mode || 'free');
  if (!fp) return { ok: false };
  const today = ymd_(new Date());
  upsertUsage_(fp, mode, today, 1);
  pruneOldUsage_();
  return { ok: true };
}

// 관리자 비밀번호 검증 → 24시간 HMAC 토큰 발급
function handleAdminAuth(p) {
  const pwd = String(p.pwd || '');
  const fp = String(p.fp || '');
  const expected = props_().getProperty('ADMIN_PASSWORD') || '';
  if (!expected) return { success: false, message: 'ADMIN_PASSWORD 미설정' };
  if (pwd !== expected) return { success: false };
  const token = makeAdminToken_(fp);
  return { success: true, token: token };
}

// 챗봇 AI 호출 (Gemini 프록시)
function handleAi(p) {
  const fp = String(p.fp || '').trim();
  const adminToken = String(p.adminToken || '');
  const prompts = Array.isArray(p.prompts) ? p.prompts : [];
  if (!prompts.length) return { error: 'no_prompts', message: '프롬프트 없음' };

  const isAdmin = adminToken && verifyAdminToken_(adminToken);
  // 챗봇 quota: 관리자 무제한, 그 외 일 30회
  const today = ymd_(new Date());
  const CHAT_LIMIT = 30;
  if (!isAdmin) {
    const u = findUsageRow_(fp, 'chat', today);
    if (u && Number(u.count || 0) >= CHAT_LIMIT)
      return { error: 'chat_limit', message: '챗봇 일일 한도 도달 (' + CHAT_LIMIT + '회)' };
  }

  const apiKey = props_().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error: 'no_api_key', message: 'GEMINI_API_KEY 미설정' };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  const results = [];
  for (let i = 0; i < prompts.length; i++) {
    const item = prompts[i] || {};
    const parts = [{ text: String(item.prompt || '') }];
    const images = Array.isArray(item.images) ? item.images : [];
    images.forEach(img => {
      if (img && img.data && img.mimeType)
        parts.push({ inline_data: { mime_type: img.mimeType, data: img.data } });
    });
    const payload = {
      contents: [{ parts: parts }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 1024 }
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
    } catch (err) {
      results.push('AI 호출 실패: ' + err.message);
    }
  }
  if (!isAdmin) upsertUsage_(fp, 'chat', today, 1);
  return { results: results };
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

// 1000건 초과 시 오래된 기록부터 자동 삭제 (개인정보처리방침 명시 사항)
function pruneOldUsage_() {
  const sh = sheet_('Usage');
  const lastRow = sh.getLastRow();
  const MAX = 1000;
  if (lastRow - 1 <= MAX) return;
  const excess = (lastRow - 1) - MAX;
  sh.deleteRows(2, excess);
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

// ───── 셋업 / 운영 도구 (편집기에서 직접 실행) ─────

// 1회 실행: 시트 + 헤더 자동 생성 + 권한 승인
function setup() {
  sheet_('Codes');
  sheet_('Usage');
  sheet_('AiLog');
  const id = props_().getProperty('SHEET_ID');
  Logger.log('✅ Setup OK. SHEET_ID=' + id);
  Logger.log('스크립트 속성에 GEMINI_API_KEY, ADMIN_PASSWORD, HMAC_SECRET 설정하세요.');
  return id;
}

// 단일 코드 발급: ADD_CODE_INPUT을 채우고 실행
function addCode() {
  const code = (ADD_CODE_INPUT.code || randCode_(6)).toUpperCase();
  const sh = sheet_('Codes');
  sh.appendRow([code, new Date(), ADD_CODE_INPUT.amount || 10000, ADD_CODE_INPUT.phone || '', '', '', 'active', ADD_CODE_INPUT.note || '']);
  Logger.log('✅ 발급된 코드: ' + code);
  return code;
}

// 미사용 코드 풀 미리 만들기 (예: bulkAddCodes(20))
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

// 코드 환불 처리 (편집기에서 실행: refundCode_apply 호출 전 REFUND_CODE 채우기)
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
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 헷갈리는 0/O/1/I 제외
  let out = '';
  for (let i = 0; i < len; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}
