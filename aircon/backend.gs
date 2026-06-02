/**
 * 청개구리 에어컨 냉매누출 진단 — GAS Backend v2.0
 * ================================================================
 * 관리자(사장님) 전용 앱. FLIR 초음파 사진 → AI(Gemini Vision) 누출 분석 +
 * 위치별 수리방법 + 분석 보고서. 점검 1건 = 시트 1행.
 *
 * ── 1회 셋업 (5분) ─────────────────────────────────────────────
 *  1. 새 Apps Script 프로젝트 생성: 이름 "청개구리-에어컨진단-Backend"
 *  2. 이 파일 전체를 Code.gs에 붙여넣기
 *  3. 스크립트 속성에 다음 추가:
 *       - GEMINI_API_KEY : 사전점검/결로진단과 같은 Gemini 키 재사용 가능 (AI 분석 필수)
 *       - APP_SECRET     : 앱 비밀번호(index.html ADMIN_PW)와 같은 값 권장 (AI 무단호출 차단)
 *                          (미설정 시 secret 검사 생략 — 키 보호 위해 설정 권장)
 *       - SHEET_ID       : (선택) 특정 시트에 기록하려면 문서 ID. 없으면 setup()이 자동 생성
 *  4. setup() 한 번 실행 → 시트 자동 생성 + SHEET_ID 자동 저장
 *  5. "배포" → "새 배포" → 유형: 웹 앱 (액세스: 모든 사용자 / 실행: 나)
 *  6. 웹 앱 URL 복사 → index.html 의 GAS_URL 상수에 붙여넣고 git push
 *
 * ── 일일 자동 리포트 (선택) ────────────────────────────────────
 *  - setupDailyTrigger() 실행 → 매일 21:00 KST sendDailyReport() 자동 호출
 *  - 텔레그램 전송하려면 스크립트 속성에 TELEGRAM_TOKEN, TELEGRAM_CHAT_ID 추가
 *    (없으면 로그에만 요약 출력)
 * ================================================================
 */

const VERSION = 'v3.0';
const SHEET_NAME = 'Inspections';
// AI 비전 모델 (사전점검 앱과 동일 키 재사용 가능)
// 사진 분석 품질 우선 → 정식 flash 사용. 최고 정확도가 필요하면 'gemini-3.1-pro',
// 비용 최소화가 필요하면 'gemini-3.1-flash-lite'로 교체
const GEMINI_MODEL = 'gemini-3.1-flash';

// 시트 헤더 (record 키 순서와 일치)
const HEADERS = [
  'savedAt','id','date','site','dong','ho','addr','inspector',
  'maker','model','refrigerant','indoorN','pipeLen','nameplateCharge',
  'leakA','leakB','leakC','leakD','leakE','leakF','leakCount','repair',
  'aiLocations','aiSummary',
  'pStart','pEnd','pHold','vac','vacRise','dtIn','dtOut','chargeActual',
  'pdrop','vacV','dtV','chgV','grade','findings','note','photoCount'
];

// ───── 진입점 ─────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'list')    return jsonRes(handleList(e.parameter));
    if (action === 'summary') return jsonRes(handleSummary(e.parameter));
    return jsonRes({ status:'ok', service:'청개구리-에어컨진단-Backend', version:VERSION });
  } catch (err) {
    return jsonRes({ error:'server_error', message:err.message });
  }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return jsonRes({ error:'bad_json', message:'잘못된 요청 형식' }); }
  const action = body.action || '';
  try {
    if (action === 'ping')            return jsonRes({ ok:true, version:VERSION });
    if (action === 'ai')              return jsonRes(handleAi(body));
    if (action === 'saveInspection')  return jsonRes(handleSave(body.record || {}));
    return jsonRes({ error:'unknown_action', valid:['ping','ai','saveInspection'] });
  } catch (err) {
    return jsonRes({ error:'server_error', message:err.message });
  }
}

// ───── 핵심: 점검 1건 저장 (id 기준 upsert) ─────
function handleSave(record) {
  if (!record || !record.id) return { ok:false, message:'record.id 없음' };
  const sh = sheet_();
  const row = HEADERS.map(h => record[h] != null ? record[h] : '');

  // 같은 id 가 이미 있으면 덮어쓰기 (재전송 멱등)
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const ids = sh.getRange(2, 2, lastRow - 1, 1).getValues(); // col B = id
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(record.id)) {
        sh.getRange(i + 2, 1, 1, HEADERS.length).setValues([row]);
        SpreadsheetApp.flush();
        return { ok:true, mode:'updated', rowNum:i + 2 };
      }
    }
  }
  sh.appendRow(row);
  SpreadsheetApp.flush();
  return { ok:true, mode:'inserted', rowNum:sh.getLastRow() };
}

// ───── AI 누출 분석 (FLIR 초음파 사진 → Gemini Vision) ─────
function handleAi(body) {
  // 간단 보호: 앱이 보낸 secret 이 APP_SECRET 과 일치해야 함 (미설정이면 검사 생략)
  const secret = props_().getProperty('APP_SECRET') || '';
  if (secret && String(body.secret || '') !== secret) return { error:'unauthorized', message:'권한 없음' };

  const apiKey = props_().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error:'no_api_key', message:'GEMINI_API_KEY 미설정 (스크립트 속성에 추가하세요)' };

  const images = Array.isArray(body.images) ? body.images : [];
  if (!images.length) return { error:'no_image', message:'사진이 없습니다' };
  if (images.length > 4) return { error:'too_many', message:'한 번에 최대 4장' };

  const ctx = body.context || {};
  const parts = [{ text: buildAiPrompt_(ctx) }];
  images.forEach(img => {
    if (img && img.data && img.mimeType) parts.push({ inline_data:{ mime_type:img.mimeType, data:img.data } });
  });

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  const payload = { contents:[{ parts:parts }], generationConfig:{ temperature:0.2, maxOutputTokens:1200 } };
  try {
    const res = UrlFetchApp.fetch(url, { method:'post', contentType:'application/json',
      payload:JSON.stringify(payload), muteHttpExceptions:true });
    if (res.getResponseCode() !== 200)
      return { error:'ai_error', message:'AI 응답 오류 (' + res.getResponseCode() + ')', raw:res.getContentText().slice(0, 300) };
    const data = JSON.parse(res.getContentText('UTF-8'));
    const txt = (data && data.candidates && data.candidates[0] && data.candidates[0].content &&
      data.candidates[0].content.parts && data.candidates[0].content.parts[0] &&
      data.candidates[0].content.parts[0].text) || '';
    return { ok:true, text:txt, parsed:tryParseJson_(txt) };
  } catch (err) {
    return { error:'ai_exception', message:err.message };
  }
}

function buildAiPrompt_(ctx) {
  const ref = ctx.refrigerant || '미상';
  const model = ctx.model || '미상';
  return [
    '당신은 에어컨 냉매누출 진단 전문가입니다.',
    '입력 사진은 FLIR 초음파(음향) 카메라로 촬영한 것으로, 냉매 누출 지점이 색상 음압(소리) 핫스팟으로 표시됩니다',
    '(보통 빨강/노랑/초록 순으로 음압이 강하며, 가장 강한 핫스팟이 누출 의심 지점입니다).',
    '아래 6개 누출 위치 중 핫스팟이 어디에 해당하는지 판단하세요:',
    'A 실내기 연결부 (플레어 너트·플레어 가공 균열/편심·단열 마감)',
    'B 실외기 연결부 (서비스밸브·캡·누유 흔적)',
    'C 분배기·용접부 (브레이징 핀홀·분배기 헤더)',
    'D 매립 배관 (구간 압력강하)',
    'E 시스템 전체 (밸브·진공·냉매 충전부)',
    'F 부가 (드레인·응축수·배관 결로)',
    '설비 정보: 냉매 ' + ref + ', 모델 ' + model + '.',
    '',
    '반드시 아래 JSON 형식 하나로만 답하세요. 코드블록·설명·서론 금지:',
    '{"leakDetected":true,"locations":[{"code":"A","name":"실내기 연결부","confidence":"높음","reason":"핫스팟 근거"}],"severity":"중대","repair":["가장 권장되는 수리방법부터 순서대로"],"summary":"2~3문장 한국어 분석 요약"}',
    'confidence 는 높음/보통/낮음, severity 는 중대/주의/양호 중 하나.',
    '핫스팟이 불명확하거나 누출 근거가 약하면 leakDetected:false, locations:[], severity:"양호" 로 답하세요.'
  ].join('\n');
}

// AI 응답에서 JSON 추출 (코드블록·앞뒤 텍스트 제거)
function tryParseJson_(txt) {
  if (!txt) return null;
  var s = String(txt).trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  var a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a >= 0 && b > a) s = s.substring(a, b + 1);
  try { return JSON.parse(s); } catch (e) { return null; }
}

// ───── 조회 ─────
function handleList(p) {
  const sh = sheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows:[] };
  const limit = Math.min(Number(p.limit) || 50, 200);
  const start = Math.max(2, lastRow - limit + 1);
  const data = sh.getRange(start, 1, lastRow - start + 1, HEADERS.length).getValues();
  const rows = data.map(r => { const o = {}; HEADERS.forEach((h, i) => o[h] = r[i]); return o; });
  return { rows: rows.reverse() };
}

function handleSummary(p) {
  const day = String(p.date || ymd_(new Date()));
  return dailyStats_(day);
}

// ───── 일일 리포트 ─────
function dailyStats_(day) {
  const sh = sheet_();
  const lastRow = sh.getLastRow();
  const stat = { date:day, total:0, 양호:0, 주의:0, 중대:0, sites:[] };
  if (lastRow < 2) return stat;
  const data = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const di = HEADERS.indexOf('date'), gi = HEADERS.indexOf('grade'),
        si = HEADERS.indexOf('site'), doi = HEADERS.indexOf('dong'), hi = HEADERS.indexOf('ho');
  data.forEach(r => {
    if (ymd_(r[di]) !== day) return;
    stat.total++;
    const g = String(r[gi] || '양호');
    if (stat[g] != null) stat[g]++;
    stat.sites.push({ site:r[si], loc:(r[doi]?r[doi]+'동 ':'')+(r[hi]?r[hi]+'호':''), grade:g });
  });
  return stat;
}

function sendDailyReport() {
  const day = ymd_(new Date());
  const s = dailyStats_(day);
  let msg = '❄️ 에어컨 냉매진단 일일 리포트 (' + day + ')\n';
  msg += '━━━━━━━━━━━━━━\n';
  msg += '총 ' + s.total + '건  ·  🟢양호 ' + s.양호 + '  🟡주의 ' + s.주의 + '  🔴중대 ' + s.중대 + '\n';
  if (s.sites.length) {
    msg += '\n';
    s.sites.forEach(x => {
      const dot = x.grade === '중대' ? '🔴' : x.grade === '주의' ? '🟡' : '🟢';
      msg += dot + ' ' + x.site + ' ' + x.loc + '\n';
    });
  }
  const token = props_().getProperty('TELEGRAM_TOKEN');
  const chat = props_().getProperty('TELEGRAM_CHAT_ID');
  if (token && chat) {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method:'post', muteHttpExceptions:true,
      payload:{ chat_id:chat, text:msg }
    });
  } else {
    Logger.log(msg);
    Logger.log('(TELEGRAM_TOKEN / TELEGRAM_CHAT_ID 미설정 — 로그 출력만)');
  }
  return s;
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(21).create();
  Logger.log('✅ 매일 21:00 일일 리포트 트리거 등록 완료');
}

// ───── 시트 헬퍼 ─────
function spreadsheet_() {
  let id = props_().getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('청개구리-에어컨진단-Data');
    id = ss.getId();
    props_().setProperty('SHEET_ID', id);
  }
  return SpreadsheetApp.openById(id);
}
function sheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ───── 유틸 ─────
function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ymd_(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return String(d);
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  return Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
}
function props_() { return PropertiesService.getScriptProperties(); }

function setup() {
  sheet_();
  const id = props_().getProperty('SHEET_ID');
  Logger.log('✅ Setup OK. SHEET_ID=' + id);
  Logger.log('웹앱 배포 후 URL 을 index.html 의 GAS_URL 에 넣으세요.');
  return id;
}
