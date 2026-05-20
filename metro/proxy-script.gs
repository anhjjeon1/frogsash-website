// ========================================
// (주)메트로 R&S AI v23.34 - Google Apps Script
// 구글시트 협업 + Drive 사진 업로드/삭제 + 행 추가/삭제 + =IMAGE() 수식 표시
// 액션: read, upload, savePhoto, migratePhotos, appendRow, deletePhoto, deleteRow, listSheets, checkCompleteColumns, addPhotoCols13, fixGunsanA1, repairBrokenRowsGunsan, inspectCell, readGrid, generateDailySalesPdf, setupDailySalesPdfTrigger, syncDashboardBeforePdf, testTelegram, setupDashboardFormulas, extendDailySalesRanges, sendDailyReportToManager, setupManagerReportTrigger, fillNoSequence, inspectPaymentSheet, setupPaymentFormulas, fixSiteTotalRanges
// v23.35: 결제현황 F열 자동집계 진단·자동화 + 사이트 합계 셀 SUM 범위 무한 확장 — inspectPaymentSheet(refCellTrace로 사이트 합계 셀 산식 추적), fixSiteTotalRanges(합계 셀 자신만 제외 + 1999행까지 확장). 진짜 정합성 누수의 근원은 각 사이트 V21 등의 `=SUM(V2:V20)` hardcoded 범위 — 합계 행 아래 새 작업 누락. 일매출 v23.25 :$X$2000 확장과 동일 철학.
// v23.34: fillNoSequence 액션 추가 — 시트별 A열(NO) 빈 셀을 마지막 NO + 1부터 연속 채움. 새 하자 행 paste 후 NO 수동 입력/드래그 채우기 자동화. B~G 컬럼 중 하나라도 데이터 있으면 채움 대상으로 인식. lastNo+1부터 시퀀셜 — 멱등 안전(이미 채워진 셀은 건드리지 않음).
// v23.33: 시스템 폴더에 '1.' prefix 적용 (매니저 공유 시 14현장만 깔끔히 다중 선택). 새 이름: '1.메트로 관리자전송', '1.메트로 당일 매출 대시보드', '1.A전체(현장관리)'. 모든 폴더 참조에 _findOrCreateSubFolder 헬퍼 도입 — 새 이름 우선 검색 → 없으면 옛 이름 fallback → 그래도 없으면 새 이름으로 생성. Drive UI 이름 변경 전·후 모두 정상 동작 (안전 전환). collectPhotosForDate SKIP 목록도 새+옛 이름 동시 등록.
// v23.32: 매니저 일일 보고 메일에 어제 사진 통합 폴더 링크 추가. 'A.메트로알엔에스(주)/메트로 관리자전송/사진_yyyy-MM-dd/{현장명}/{동-호}/{수리전|수리후|확인서}.{ext}' 계층 구조로 어제 사진 사본 생성 후 ANYONE_WITH_LINK+VIEW 공유 → 폴더 링크 1개 클릭 → Drive 우상단 다운로드로 zip 한 방. 현장명·동호 폴더 트리로 자동 명기. 같은 날짜 재실행 시 휴지통 후 재생성(멱등). 원본 동호 폴더는 보존(사본만 추가).
// v23.31: 메트로 관리자(구미영 대리, era999@naver.com)에게 매일 09:00 KST 일일 보고 메일 자동 발송. 어제 작업 사진(Drive 보기 링크) + claude현장관리(종합)_LIVE xlsx 사본(일매출 시트 제외) 첨부. xlsx는 'A.메트로알엔에스(주)/메트로 관리자전송/'에 일별 보관(감사 추적). 받은 사람은 xlsx 다운받아 자유 편집 가능하나 본사 원본 시트엔 영향 없음. Script Properties MANAGER_EMAIL 필요.
// v23.30: savePhoto/migratePhotos에서 _data 열 저장 값을 base64 → Drive URL 로 변경. Google Sheets 셀당 50000자 한계로 인해 1280px·0.85 사진 base64(~200KB+ 텍스트)가 setValue throw되며 응답 실패하던 사고 영구 해결. 이미지 열은 이미 =IMAGE(URL) 수식이라 변경 무관, _data 열도 read 액션이 'http' 시작 매칭으로 그대로 처리. Drive 폴더 구조 (A.메트로알엔에스(주)/{현장}/{동}-{호}/) 그대로 유지. 결과: 사진 크기 무제한, 시트 용량 부담 급감, 클라이언트 1920px·0.92 고화질 풀 동작.
// v23.25: 일매출 시트 SUMPRODUCT 범위 확장 — extendDailySalesRanges 액션 추가. K/L/M/O 컬럼 :$X$NNN → :$X$2000 일괄 치환(시작 :$X$2 보호 + 단가표 Q/U 자동 보호). 시트별 hardcoded 행수(동탄 138, 광주중흥 154, 양주 N 등)가 시트 확장 페이스를 못 따라가던 정합성 누수 영구 해결. 5/8 동탄 25건 완료(NO 139~)가 일매출 0원으로 떨어졌던 사고 재발 방지.
// v23.24: 입출금 현황 요약 단일 진본화 — 대시보드 M12~M17을 결제현황 시트 자동 참조 수식으로 전환. syncDashboardBeforePdf는 더 이상 setValue로 덮어쓰지 않고 셀에서 계산된 값을 읽어 PDF·텔레그램에 사용. 신규 액션 setupDashboardFormulas로 셀 수식 6개 + 월별 매출 추이 차트(2026년만) 1회 셋업. 입금 횟수 R29:R1000 동적 카운트(R46 누락 버그 수정).
// v23.23: 텔레그램 일일 보고 알림 추가 — generateDailySalesPdf 직후 비공개 채널에 요약+PDF링크 푸시. Script Properties (TG_TOKEN, TG_CHAT_ID) 필요. 미설정 시 알림만 건너뜀, PDF 생성은 정상.
// v23.22: A3 헤더에 당일 매출 추가 — '📅 기준일: yyyy년 MM월 dd일   💰 당일 매출: N,NNN,NNN원'. financial M12·M14는 정수 반올림(부동소수점 .36 표시 제거).
// v23.21: PDF 생성 직전 대시보드 시트 자동 동기화 — 기준일(A3) 갱신 + financial 영역(M12~M17)을 일매출 시트의 오늘 미지급 + 결제현황 입금 합계 기반으로 자동 갱신. 일매출 시트와 PDF의 미지급 잔액 항상 일치, 입금 횟수 자동 카운트.
// v23.20: 일매출 대시보드 PDF 자동 생성 — _LIVE 대시보드 시트를 PDF로 export 후 Drive 'A.메트로알엔에스(주)/메트로 당일 매출 대시보드/' 저장. 매일 21:30 KST 시간 트리거 (폴더 NFC 정규화 + 옛 이름 fallback + 자동 생성)
// v23.19: readGrid 액션 추가 — 시트의 raw 2D 배열 그대로 반환 (METRO-APP/calendar_sync가 xlsm 대신 _LIVE 직접 사용)
//         xlsm 머지 손상 사고 후 데이터 진본을 _LIVE 단일 관리로 전환하기 위한 핵심 API
// v23.18: inspectCell 진단 액션 추가 — 특정 행(row 또는 NO 검색)의 사진 H/I/J + _data 컬럼 formula·value 일괄 반환
//         M5.5 2단계 사진 IMAGE 수식 누락 원인 진단용 (savePhoto가 setFormula 호출했는지 확인)
// v23.17: read 응답 _photos를 URL 우선으로 (IMAGE 수식의 Drive URL 1순위, base64 2순위)
//         → Python xlsm 머지가 IMAGE 수식 작성 가능. 메트로앱은 URL/base64 둘 다 img src로 동작
// v23.16: 군산미장 A1 헤더 'ㅡDUF'→'NO' 영구 수정 + NO 비어있는 깨진 행 일괄 NO 채번·서식 복사
// v23.15: 13시트(군산미장 제외) H~J 사진 컬럼 일괄 추가 함수 + HTTP 액션 — M4.5 양식 통일
// v23.11: appendRow NO 채번 + 서식 복사 거꾸로 스캔 — lastRow가 빈 양식이어도 정상 행을 찾아 적용
//         (이전 행이 깨져 있어도 새 행은 정상 양식으로 들어감 — 체인 깨짐 방지)
// v23.10: appendRow NO 채번 A열 fallback — 헤더 매칭 실패 시 A열 마지막 값이 숫자면 A열을 NO로 가정
//         (군산미장 A1='ㅡDUF' 같은 깨진 헤더에도 NO 자동 채번 동작하도록)
// v23.9: appendRow에 이전 행 서식 복사 추가 (PASTE_FORMAT) — 새 하자 행 테두리·정렬 자동 적용
// v23.8: savePhoto 자동완료 — L(완료) 컬럼은 '완료'가 아닌 모든 값을 '완료'로 덮어쓰기
//        (기존 '미완료'/'N' 값이 그대로 남아 토글이 미완료로 인식하던 문제 수정)
// v23.7: read API 객체 변환 첫 매칭 우선 — R열 통계 '완료'가 L열 행별 '완료'를 덮어쓰는 버그 수정
//        클라이언트 완료 행 숨김 토글이 정상 동작하도록
// v23.6: savePhoto 자동완료 + 진단 함수 모두 첫 매칭(가장 좌측) 우선 — 통계 표 '완료' 헤더 오인 버그 수정
//        진단 함수가 모든 매칭 위치를 리스트로 반환 (행별 컬럼 vs 통계 헤더 구분)
// v23.5: doGet에 checkCompleteColumns 액션 추가 — HTTP로 14시트 K/L 헤더 위치 진단 (재배포 1회 후 자동 호출 가능)
// v23.4: M4 검증용 진단 함수 oneTimeCheckCompleteColumns 추가 — 14시트 K(완료일)/L(완료) 헤더 위치 일괄 점검
// v23.2: 14개 시트 H~J 사진 컬럼(수리전/수리후/완료확인서) 일괄 추가 — oneTimeAddPhotoColumnsToAllSites
//        + savePhoto 안전망: 사진 컬럼 없으면 H~J 자동 삽입
// v23.1(M3): 사진 저장 경로 변경 — A.메트로알엔에스(주)/{현장}/{동}-{호}/ (동·호 단위 분리)
//           일회성 정리 함수 oneTimeOrganizeDriveFolders 추가
// v23.0: listSheets 액션 추가 — 워크북의 현장 시트(14개) 동적 로딩
// v21.10: 사진 업로드 시 행 높이 + 컬럼 너비 모두 160px
// ========================================

// 사진 저장 루트 — Drive 내 폴더 이름 (v23.1)
var DRIVE_PHOTO_ROOT = 'A.메트로알엔에스(주)';

// 시스템 시트 (드롭다운에서 제외) — 현장 시트는 이 목록 외 전체
var SYSTEM_SHEETS = {
  '대시보드': true, '일매출': true, '전체공정표': true,
  '결제현황': true, '단가표': true,
  '2025년매출': true, '2026년매출': true
};

// 권한 재승인 트리거용 — GAS 편집기에서 직접 실행
function authorizeAll() {
  try {
    // SpreadsheetApp 접근
    var files = DriveApp.getFilesByName('__metro_auth_test__');
    while (files.hasNext()) files.next();
    // Drive 쓰기 권한
    var f = DriveApp.createFile('__metro_auth_test__.txt', 'auth test', 'text/plain');
    f.setTrashed(true);
    Logger.log('✅ 권한 승인 완료: SpreadsheetApp + DriveApp');
    return 'OK';
  } catch(e) {
    Logger.log('❌ 권한 오류: ' + e.message);
    throw e;
  }
}

function makeRes(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Drive 헬퍼: 자식 폴더 확보 (없으면 생성)
function _getOrCreateSubFolder(parent, name) {
  var subs = parent.getFoldersByName(name);
  return subs.hasNext() ? subs.next() : parent.createFolder(name);
}

// [v23.33] 자식 폴더 확보 — 새 이름 우선, 옛 이름 fallback, 없으면 새 이름으로 생성
// Drive UI에서 이름 변경 전·후 모두 정상 동작 (안전 전환)
// 사용 예: _findOrCreateSubFolder(rootFolder, '1.메트로 관리자전송', ['메트로 관리자전송'])
function _findOrCreateSubFolder(parent, primaryName, aliases) {
  var iter = parent.getFoldersByName(primaryName);
  if (iter.hasNext()) return iter.next();
  if (aliases && aliases.length) {
    for (var i = 0; i < aliases.length; i++) {
      var fb = parent.getFoldersByName(aliases[i]);
      if (fb.hasNext()) return fb.next();
    }
  }
  return parent.createFolder(primaryName);
}

// Drive 사진 저장 루트 폴더 가져오기 (v23.1: A.메트로알엔에스(주))
function _getPhotoRoot() {
  var roots = DriveApp.getFoldersByName(DRIVE_PHOTO_ROOT);
  if (!roots.hasNext()) {
    // 없으면 새로 생성 (My Drive 직속)
    return DriveApp.createFolder(DRIVE_PHOTO_ROOT);
  }
  return roots.next();
}

// v23.1: 동·호 있으면 새 경로 (A.메트로알엔에스(주)/{현장}/{동}-{호}/), 없으면 fallback (기존 METRO_PHOTOS)
function getOrCreatePhotoFolder(sheetId, sheetName, dong, ho) {
  var d = String(dong || '').trim();
  var h = String(ho || '').trim();
  if (d && h) {
    var root = _getPhotoRoot();
    var siteFolder = _getOrCreateSubFolder(root, String(sheetName || 'sheet'));
    return _getOrCreateSubFolder(siteFolder, d + '-' + h);
  }
  // fallback (동·호 없을 때) — v23.0 이하 경로 유지
  var rootName = 'METRO_PHOTOS';
  var roots = DriveApp.getFoldersByName(rootName);
  var fbRoot = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
  var subName = (sheetName || 'sheet') + '_' + String(sheetId).substring(0, 8);
  var subs = fbRoot.getFoldersByName(subName);
  return subs.hasNext() ? subs.next() : fbRoot.createFolder(subName);
}

// base64 → Drive 업로드 + 공개 공유
function uploadPhotoToDrive(base64, sheetId, sheetName, rowNum, photoType, dong, ho) {
  var match = base64.match(/data:(.*?);base64,(.*)/);
  if (!match) throw new Error('잘못된 base64 데이터');
  var mime = match[1];
  var bytes = Utilities.base64Decode(match[2]);
  var ext = (mime.split('/')[1] || 'jpg').split('+')[0];
  var fname = 'row' + rowNum + '_' + photoType + '_' + new Date().getTime() + '.' + ext;
  var blob = Utilities.newBlob(bytes, mime, fname);
  var folder = getOrCreatePhotoFolder(sheetId, sheetName, dong, ho);
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return {
    id: file.getId(),
    url: 'https://lh3.googleusercontent.com/d/' + file.getId()
  };
}

// === [v23.1 일회성 정리] Drive 폴더 시트명으로 통일 + 불필요 폴더 휴지통 ===
// 실행 방법: GAS 편집기 좌상단 함수 선택 → oneTimeOrganizeDriveFolders → ▶ 실행
// 한 번만 실행하면 됨 (멱등 — 재실행해도 부작용 없음)
function oneTimeOrganizeDriveFolders() {
  var roots = DriveApp.getFoldersByName(DRIVE_PHOTO_ROOT);
  if (!roots.hasNext()) {
    Logger.log('❌ 루트 폴더 없음: ' + DRIVE_PHOTO_ROOT);
    return {error:'root not found'};
  }
  var root = roots.next();

  var renameMap = {
    '감일제일건설': '감일제일',
    '검단제일건설': '검단제일',
    '경산하양제일건설': '경산하양',
    '광주중흥제일건설': '광주중흥',
    '군산제일건설': '군산미장',
    '동탄제일건설(41단지)': '동탄',
    '양산제일건설': '양산',
    '양주제일건설(1단지)': '양주',
    '원주(무실)제일건설': '원주(무실)',
    '원주(혁신)제일건설': '원주(혁신)',
    '익산제일건설': '익산제일',
    '충주호암제일건설': '충주호암',
    '파주제일건설(1단지)': '파주1단지',
    '파주제일제일(6단지)': '파주6단지'
  };
  var trashList = ['하자전후(루버편)'];

  var renamed = [], trashed = [], skipped = [];
  var sub = root.getFolders();
  while (sub.hasNext()) {
    var f = sub.next();
    var nm = f.getName();
    if (renameMap[nm]) {
      f.setName(renameMap[nm]);
      Logger.log('✏️ rename: ' + nm + ' → ' + renameMap[nm]);
      renamed.push(nm + ' → ' + renameMap[nm]);
    } else if (trashList.indexOf(nm) >= 0) {
      f.setTrashed(true);
      Logger.log('🗑️ trash: ' + nm);
      trashed.push(nm);
    } else {
      skipped.push(nm);
    }
  }
  var result = {renamed:renamed, trashed:trashed, skipped:skipped,
    summary: 'rename=' + renamed.length + ', trash=' + trashed.length + ', skip=' + skipped.length};
  Logger.log('=== 완료 — ' + result.summary);
  Logger.log('그대로 둔 폴더: ' + skipped.join(', '));
  return result;
}

// === [v23.2 일회성] 14개 시트 H~J에 사진 컬럼(수리전/수리후/완료확인서) 일괄 추가 ===
// 실행: GAS 편집기에서 oneTimeAddPhotoColumnsToAllSites 직접 실행
// 멱등 — 이미 있는 시트는 자동 스킵, 재실행 안전
function oneTimeAddPhotoColumnsToAllSites() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';  // _LIVE
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var allSheets = ss.getSheets();

  var processed = [], skipped = [], errored = [];
  var photoCols = ['수리전', '수리후', '완료확인서'];
  var dataCols = ['수리전_data', '수리후_data', '완료확인서_data'];

  for (var s = 0; s < allSheets.length; s++) {
    var ws = allSheets[s];
    var name = ws.getName();
    if (SYSTEM_SHEETS[name]) { skipped.push(name + ' (시스템)'); continue; }

    try {
      var lastCol = ws.getLastColumn();
      if (lastCol < 1) { skipped.push(name + ' (빈 시트)'); continue; }
      var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
      var hasPhoto = false;
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        if (hn === '수리전' || hn === '수리후' || hn === '완료확인서' || hn === '확인서') {
          hasPhoto = true; break;
        }
      }
      if (hasPhoto) { skipped.push(name + ' (이미 있음)'); continue; }

      // ① H 위치(8번째)에 3개 컬럼 삽입 → H/I/J = 수리전/수리후/완료확인서
      ws.insertColumnsBefore(8, 3);
      for (var p = 0; p < 3; p++) {
        ws.getRange(1, 8 + p).setValue(photoCols[p]).setFontWeight('bold');
        try { ws.setColumnWidth(8 + p, 160); } catch(e) {}
      }

      // ② 시트 끝에 _data 3개 추가 (숨김)
      var endCol = ws.getLastColumn();
      for (var d = 0; d < 3; d++) {
        ws.getRange(1, endCol + 1 + d).setValue(dataCols[d]).setFontWeight('bold');
      }
      ws.hideColumns(endCol + 1, 3);

      processed.push(name);
      Logger.log('✅ ' + name + ': H~J 본 컬럼 + 끝 _data 3개 추가');
    } catch(e) {
      errored.push(name + ': ' + e.message);
      Logger.log('❌ ' + name + ' 오류: ' + e.message);
    }
  }

  SpreadsheetApp.flush();
  Logger.log('=== 완료 — 처리: ' + processed.length + ', 스킵: ' + skipped.length + ', 에러: ' + errored.length);
  Logger.log('처리: ' + processed.join(', '));
  Logger.log('스킵: ' + skipped.join(', '));
  if (errored.length) Logger.log('에러: ' + errored.join('; '));
  return {processed:processed, skipped:skipped, errored:errored};
}

// === [v23.16 일회성] 군산미장 A1 헤더 'ㅡDUF' → 'NO' 영구 수정 ===
function oneTimeFixGunsanA1() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('군산미장');
  if (!ws) return {error: '군산미장 시트 없음'};
  var current = String(ws.getRange(1, 1).getValue() || '').trim();
  if (current === 'NO') return {status: 'already_NO', before: current};
  ws.getRange(1, 1).setValue('NO').setFontWeight('bold');
  SpreadsheetApp.flush();
  return {status: 'ok', before: current, after: 'NO'};
}

// === [v23.16 일회성] 군산미장에서 NO 비어 있고 데이터 있는 행 일괄 정리 ===
// NO 자동 채번(max+1) + 정상 행의 서식 복사
function oneTimeRepairBrokenRowsInGunsan() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('군산미장');
  if (!ws) return {error: '군산미장 시트 없음'};

  var lastRow = ws.getLastRow();
  var lastCol = ws.getLastColumn();

  // 헤더에서 NO 컬럼 찾기 (v23.16 후엔 A열='NO')
  var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  var colNo = -1, colDong = -1, colHo = -1;
  for (var h = 0; h < headers.length; h++) {
    var hn = String(headers[h] || '').replace(/\s/g,'');
    if (hn === 'NO' || hn === '번호') colNo = h + 1;
    if (hn === '동') colDong = h + 1;
    if ((hn === '호수' || hn === '호') && colHo < 0) colHo = h + 1;
  }
  if (colNo < 0 && lastRow >= 2) {
    // fallback: A열에 숫자 있는 행이 있으면 A열을 NO로
    for (var fc = lastRow; fc >= 2; fc--) {
      var v = String(ws.getRange(fc, 1).getValue() || '').trim();
      if (/^\d+$/.test(v)) { colNo = 1; break; }
    }
  }
  if (colNo < 0) return {error: 'NO 컬럼 못 찾음'};

  // 1) max NO + 정상 행 (서식 source) 찾기
  var maxNo = 0;
  var srcFmtRow = -1;
  var noVals = ws.getRange(2, colNo, lastRow - 1, 1).getValues();
  for (var i = 0; i < noVals.length; i++) {
    var n = parseInt(noVals[i][0], 10);
    if (!isNaN(n)) {
      if (n > maxNo) maxNo = n;
      srcFmtRow = i + 2; // 가장 마지막 정상 행
    }
  }

  // 2) NO 비어 있고 데이터 있는 행 찾기 + 정리
  var fixed = [];
  for (var r = 2; r <= lastRow; r++) {
    var noVal = String(ws.getRange(r, colNo).getValue() || '').trim();
    if (noVal) continue;
    var hasData = false;
    if (colDong > 0 && String(ws.getRange(r, colDong).getValue() || '').trim()) hasData = true;
    if (!hasData && colHo > 0 && String(ws.getRange(r, colHo).getValue() || '').trim()) hasData = true;
    if (!hasData) continue;
    // NO 채번
    maxNo++;
    ws.getRange(r, colNo).setValue(maxNo);
    // 서식 복사 (정상 행 1개만 source)
    if (srcFmtRow > 0 && srcFmtRow !== r) {
      try {
        ws.getRange(srcFmtRow, 1, 1, lastCol).copyTo(
          ws.getRange(r, 1, 1, lastCol),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
      } catch(e) {}
    }
    fixed.push({row: r, no: maxNo});
  }

  SpreadsheetApp.flush();
  return {status: 'ok', fixedCount: fixed.length, fixed: fixed};
}

// === [v23.15 일회성] 13시트(군산미장 제외)에 H~J 사진 컬럼 추가 ===
// 군산미장은 v23.2에서 이미 처리됨. 13시트만 양식 통일 (옵션 2: H 위치에 삽입)
// 멱등 — 이미 사진 컬럼 있는 시트는 자동 스킵
function oneTimeAddPhotoColumnsTo13Sheets() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var allSheets = ss.getSheets();

  var EXEMPT = {'군산미장': true}; // 이미 H~J 사진 컬럼 있음
  var processed = [], skipped = [], errored = [];
  var photoCols = ['수리전', '수리후', '완료확인서'];
  var dataCols = ['수리전_data', '수리후_data', '완료확인서_data'];

  for (var s = 0; s < allSheets.length; s++) {
    var ws = allSheets[s];
    var name = ws.getName();
    if (SYSTEM_SHEETS[name]) { skipped.push(name + ' (시스템)'); continue; }
    if (EXEMPT[name]) { skipped.push(name + ' (이미 적용됨)'); continue; }

    try {
      var lastCol = ws.getLastColumn();
      if (lastCol < 1) { skipped.push(name + ' (빈 시트)'); continue; }
      var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];

      // 정확 매칭으로 사진 컬럼 존재 여부 체크
      var hasPhoto = false;
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        if (hn === '수리전' || hn === '수리후' || hn === '완료확인서') { hasPhoto = true; break; }
      }
      if (hasPhoto) { skipped.push(name + ' (이미 사진 컬럼)'); continue; }

      // ① H 위치(8번째)에 3개 컬럼 삽입
      ws.insertColumnsBefore(8, 3);
      for (var p = 0; p < 3; p++) {
        ws.getRange(1, 8 + p).setValue(photoCols[p]).setFontWeight('bold');
        try { ws.setColumnWidth(8 + p, 160); } catch(e) {}
      }

      // ② 시트 끝에 _data 3개 (숨김)
      var endCol = ws.getLastColumn();
      for (var d = 0; d < 3; d++) {
        ws.getRange(1, endCol + 1 + d).setValue(dataCols[d]).setFontWeight('bold');
      }
      try { ws.hideColumns(endCol + 1, 3); } catch(e) {}

      processed.push(name);
      Logger.log('✅ ' + name);
    } catch(e) {
      errored.push(name + ': ' + e.message);
      Logger.log('❌ ' + name + ': ' + e.message);
    }
  }

  SpreadsheetApp.flush();
  var summary = 'processed=' + processed.length + ', skipped=' + skipped.length + ', errored=' + errored.length;
  Logger.log('=== ' + summary);
  return {processed: processed, skipped: skipped, errored: errored, summary: summary};
}

// === [v23.6 진단] 14시트의 "완료일"/"완료" 헤더 위치 일괄 점검 ===
// 실행: GAS 편집기에서 oneTimeCheckCompleteColumns 직접 실행 또는 ?action=checkCompleteColumns
// v23.6: 모든 매칭 위치를 리스트로 반환 (통계 표의 '완료' 헤더 vs 행별 '완료' 구분)
//        + 첫 매칭(가장 좌측)을 진짜 행별 컬럼으로 판정
function oneTimeCheckCompleteColumns() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var allSheets = ss.getSheets();

  var sheets = [];

  for (var s = 0; s < allSheets.length; s++) {
    var ws = allSheets[s];
    var name = ws.getName();
    if (SYSTEM_SHEETS[name]) continue;

    var lastCol = ws.getLastColumn();
    if (lastCol < 1) { sheets.push({name:name, error:'빈 시트'}); continue; }
    var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];

    var doneDateAll = [], doneAll = [];
    for (var h = 0; h < headers.length; h++) {
      var hn = String(headers[h]).replace(/\s/g,'');
      if (hn === '완료일') doneDateAll.push(h + 1);
      else if (hn === '완료') doneAll.push(h + 1);
    }

    // 첫 매칭 = 행별 컬럼 (좌측이 데이터, 우측이 통계 표 헤더라는 가정)
    var doneDateFirst = doneDateAll.length ? doneDateAll[0] : -1;
    var doneFirst = doneAll.length ? doneAll[0] : -1;

    sheets.push({
      name: name,
      doneDate: doneDateFirst > 0 ? _colLetter(doneDateFirst) : null,
      done: doneFirst > 0 ? _colLetter(doneFirst) : null,
      doneDateAll: doneDateAll.map(_colLetter),
      doneAll: doneAll.map(_colLetter),
      hasDup: doneDateAll.length > 1 || doneAll.length > 1
    });
  }

  Logger.log('=== M4 검증 v23.6: 완료일/완료 컬럼 점검 ===');
  sheets.forEach(function(s){
    if (s.error) { Logger.log('❌ ' + s.name + ': ' + s.error); return; }
    var dup = s.hasDup ? ' [중복!]' : '';
    Logger.log(s.name + ': 완료일=' + s.doneDate + ' / 완료=' + s.done +
      (s.hasDup ? ' (완료일 후보=' + s.doneDateAll.join(',') + ', 완료 후보=' + s.doneAll.join(',') + ')' : '') + dup);
  });
  return {sheets: sheets, count: sheets.length};
}

// 컬럼 인덱스(1-based) → 알파벳
function _colLetter(idx) {
  var s = '';
  while (idx > 0) {
    var r = (idx - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

// 기존 이미지 셀의 =IMAGE("...") 수식에서 Drive 파일 ID 추출 후 삭제
function tryTrashOldImageFile(ws, rowNum, imgColIdx) {
  try {
    var f = ws.getRange(rowNum, imgColIdx).getFormula();
    if (!f) return;
    var m = f.match(/\/d\/([A-Za-z0-9_-]+)/) || f.match(/id=([A-Za-z0-9_-]+)/);
    if (!m) return;
    DriveApp.getFileById(m[1]).setTrashed(true);
  } catch(e) {}
}

// === GET 요청 ===
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'read') {
    var sheetId = e.parameter.sheetId;
    var sheetName = e.parameter.sheetName || '';
    if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});
    try {
      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음: '+sheetName});

      var lastRow = ws.getLastRow();
      var lastCol = ws.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return makeRes({status:'ok', rows:[], sheetName:ws.getName(), count:0});

      var data = ws.getRange(1, 1, lastRow, lastCol).getValues();
      var formulas = ws.getRange(1, 1, lastRow, lastCol).getFormulas();
      var headers = data[0];

      // _data 열 (base64 텍스트 저장)
      var photoDataCols = {};
      // 이미지 열 (CellImage/IMAGE 수식)
      var photoImgCols = {};
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        if (hn === '수리전_data') photoDataCols.before = h;
        else if (hn === '수리후_data') photoDataCols.after = h;
        else if (hn === '완료확인서_data' || hn === '확인서_data') photoDataCols.confirm = h;
        else if (hn === '수리전') photoImgCols.before = h;
        else if (hn === '수리후') photoImgCols.after = h;
        else if (hn === '완료확인서' || hn === '확인서') photoImgCols.confirm = h;
      }

      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          // v23.7: 첫 매칭 우선 — 같은 헤더 이름이 두 번 나오면(예: 행별 '완료' L열 + 통계 '완료' R열)
          //         첫 번째(좌측) 값을 보존. 그래야 클라이언트가 행별 '완료'를 정확히 읽음
          if (Object.prototype.hasOwnProperty.call(obj, headers[j])) continue;
          var isPhotoCol = false;
          for (var pt in photoImgCols) { if (photoImgCols[pt] === j) { isPhotoCol = true; break; } }
          for (var pt2 in photoDataCols) { if (photoDataCols[pt2] === j) { isPhotoCol = true; break; } }
          if (isPhotoCol) {
            obj[headers[j]] = '';
          } else {
            obj[headers[j]] = data[i][j] !== undefined ? String(data[i][j]) : '';
          }
        }
        obj._rowNum = i + 1;

        obj._photos = {};

        // v23.17: URL 우선 — Excel xlsm IMAGE 수식과 호환되도록 IMAGE() 수식 URL을 1순위로
        // 1순위: 이미지 열 IMAGE 수식의 URL (Drive 공개 URL)
        for (var pType in photoImgCols) {
          var colI = photoImgCols[pType];
          var fm = (formulas[i] && formulas[i][colI]) ? String(formulas[i][colI]) : '';
          if (fm) {
            var um = fm.match(/"(https?:\/\/[^"]+)"/);
            if (um) { obj._photos[pType] = um[1]; continue; }
          }
        }

        // 2순위: _data 열 base64 (URL이 없는 행만)
        for (var pType2 in photoDataCols) {
          if (obj._photos[pType2]) continue;
          var col2 = photoDataCols[pType2];
          var val = String(data[i][col2] || '');
          if (val.indexOf('data:image') === 0 || val.indexOf('http') === 0) {
            obj._photos[pType2] = val;
          }
        }

        // 3순위: 이미지 열의 raw 값 (구 데이터, 수식 없이 base64나 URL이 직접 들어있는 경우)
        for (var pType3 in photoImgCols) {
          if (obj._photos[pType3]) continue;
          var colI3 = photoImgCols[pType3];
          var raw = String(data[i][colI3] || '');
          if (raw.indexOf('data:image') === 0 || raw.indexOf('http') === 0) {
            obj._photos[pType3] = raw;
          }
        }

        rows.push(obj);
      }
      return makeRes({status:'ok', rows:rows, sheetName:ws.getName(), count:rows.length});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === 시트 목록 (현장 드롭다운용, v23.0) ===
  if (action === 'listSheets') {
    var sheetId = e.parameter.sheetId;
    if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});
    try {
      var ss = SpreadsheetApp.openById(sheetId);
      var all = ss.getSheets();
      var sites = [];
      for (var s = 0; s < all.length; s++) {
        var nm = all[s].getName();
        if (SYSTEM_SHEETS[nm]) continue;
        // 행 수 (헤더 제외 추정치)
        var lr = all[s].getLastRow();
        sites.push({name: nm, rowCount: Math.max(0, lr - 1)});
      }
      return makeRes({status:'ok', sites:sites, count:sites.length});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.16] 군산미장 A1 헤더 ㅡDUF → NO ===
  if (action === 'fixGunsanA1') {
    try {
      var r1 = oneTimeFixGunsanA1();
      return makeRes(Object.assign({status:'ok'}, r1));
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.16] 군산미장 NO 비어있는 행 일괄 정리 ===
  if (action === 'repairBrokenRowsGunsan') {
    try {
      var r2 = oneTimeRepairBrokenRowsInGunsan();
      return makeRes(Object.assign({status:'ok'}, r2));
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.15] 13시트 H~J 사진 컬럼 일괄 추가 (M4.5, HTTP 호출 가능) ===
  if (action === 'addPhotoCols13') {
    try {
      var result = oneTimeAddPhotoColumnsTo13Sheets();
      return makeRes({
        status:'ok',
        processed: result.processed,
        skipped: result.skipped,
        errored: result.errored,
        summary: result.summary
      });
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.5+] 14시트 완료일/완료 컬럼 진단 (M4 검증용, HTTP 호출 가능) ===
  if (action === 'checkCompleteColumns') {
    try {
      var result = oneTimeCheckCompleteColumns();
      return makeRes({
        status:'ok',
        sheets: result.sheets,
        count: result.count
      });
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.18] 사진 IMAGE 수식 진단 — M5.5 2단계 ===
  // 사용 예: ?action=inspectCell&sheetId=...&sheetName=경산하양&no=108
  //         ?action=inspectCell&sheetId=...&sheetName=경산하양&row=110
  if (action === 'inspectCell') {
    var sheetId = e.parameter.sheetId;
    var sheetName = e.parameter.sheetName || '';
    var rowParam = e.parameter.row;
    var noParam = e.parameter.no;
    if (!sheetId || !sheetName) return makeRes({status:'error', message:'sheetId, sheetName 필요'});
    try {
      var ss = SpreadsheetApp.openById(sheetId);
      var ws = ss.getSheetByName(sheetName);
      if (!ws) return makeRes({status:'error', message:'시트 없음: '+sheetName});

      var lastCol = ws.getLastColumn();
      var lastRow = ws.getLastRow();
      var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];

      // NO 컬럼(첫 매칭) 찾기
      var noCol = -1;
      for (var hi = 0; hi < headers.length; hi++) {
        if (String(headers[hi]).replace(/\s/g,'') === 'NO') { noCol = hi + 1; break; }
      }

      var rowNum = parseInt(rowParam || '0');
      // NO 값으로 행 검색
      if (!rowNum && noParam !== undefined && String(noParam) !== '') {
        if (noCol < 0) return makeRes({status:'error', message:'NO 헤더를 찾을 수 없음'});
        var noVals = ws.getRange(2, noCol, Math.max(1, lastRow - 1), 1).getValues();
        for (var ri = 0; ri < noVals.length; ri++) {
          if (String(noVals[ri][0]) === String(noParam)) { rowNum = ri + 2; break; }
        }
        if (!rowNum) return makeRes({status:'error', message:'NO '+noParam+' 행을 찾을 수 없음'});
      }
      if (!rowNum || rowNum < 2 || rowNum > lastRow) {
        return makeRes({status:'error', message:'유효하지 않은 row: '+rowNum+' (lastRow='+lastRow+')'});
      }

      // 사진 컬럼 매핑 (첫 매칭 우선)
      var photoMap = {};
      var photoAliases = ['수리전','수리후','완료확인서','확인서','수리전_data','수리후_data','완료확인서_data','확인서_data'];
      for (var hh = 0; hh < headers.length; hh++) {
        var hnh = String(headers[hh]).replace(/\s/g,'');
        for (var ai = 0; ai < photoAliases.length; ai++) {
          if (hnh === photoAliases[ai] && !photoMap[photoAliases[ai]]) {
            photoMap[photoAliases[ai]] = hh + 1;
          }
        }
      }

      var rowValues = ws.getRange(rowNum, 1, 1, lastCol).getValues()[0];
      var rowFormulas = ws.getRange(rowNum, 1, 1, lastCol).getFormulas()[0];

      // 사진 셀 요약
      var photoSummary = {};
      for (var pa = 0; pa < photoAliases.length; pa++) {
        var key = photoAliases[pa];
        if (!photoMap[key]) continue;
        var col = photoMap[key];
        var ci = col - 1;
        var v = String(rowValues[ci] || '');
        var f = String(rowFormulas[ci] || '');
        var entry = {
          col: col,
          a1: ws.getRange(rowNum, col).getA1Notation(),
          formula: f,
          hasImageFormula: f.indexOf('IMAGE(') >= 0,
          valueLength: v.length,
          valuePreview: v.length > 80 ? v.substring(0, 80) + '...' : v
        };
        if (key.indexOf('_data') >= 0) {
          entry.isBase64 = v.indexOf('data:image') === 0;
        }
        photoSummary[key] = entry;
      }

      // 핵심 헤더 요약
      var rowSummary = {};
      var keyHeaders = ['NO','동','호','호수','위치','하자내용','완료일','완료'];
      for (var ki = 0; ki < keyHeaders.length; ki++) {
        for (var hi2 = 0; hi2 < headers.length; hi2++) {
          var hnk = String(headers[hi2]).replace(/\s/g,'');
          if (hnk === keyHeaders[ki] && !Object.prototype.hasOwnProperty.call(rowSummary, keyHeaders[ki])) {
            rowSummary[keyHeaders[ki]] = String(rowValues[hi2] || '');
            break;
          }
        }
      }

      return makeRes({
        status:'ok',
        sheetName: ws.getName(),
        rowNum: rowNum,
        lastRow: lastRow,
        lastCol: lastCol,
        rowSummary: rowSummary,
        photoColumns: photoMap,
        photoSummary: photoSummary
      });
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.19] readGrid — 시트의 raw 2D 배열 그대로 반환 (METRO-APP/calendar_sync용) ===
  // 사용 예: ?action=readGrid&sheetId=...&sheetName=대시보드
  //         ?action=readGrid&sheetId=...&sheetName=*  (모든 시스템 시트 + 14현장 한 번에)
  if (action === 'readGrid') {
    var sheetId = e.parameter.sheetId;
    var sheetName = e.parameter.sheetName || '';
    if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});
    try {
      var ss = SpreadsheetApp.openById(sheetId);

      function gridOf(ws) {
        var lr = ws.getLastRow();
        var lc = ws.getLastColumn();
        if (lr < 1 || lc < 1) return {name: ws.getName(), rows: 0, cols: 0, values: []};
        // getDisplayValues로 가져오면 날짜·통화 형식이 표시값으로 그대로 옴 — 파이썬에서 cell()이 처리하기 쉬움
        // 단 숫자 계산이 필요하면 getValues()로 받아야. _data 컬럼(base64) 통신 부담 줄이기 위해 H/I/J·_data는 비움
        var vals = ws.getRange(1, 1, lr, lc).getValues();
        var headers = vals[0];
        var skipCols = {};
        for (var h = 0; h < headers.length; h++) {
          var hn = String(headers[h]).replace(/\s/g,'');
          // 사진 컬럼·_data는 통신 부담 큼 — 빈 값으로 치환 (METRO-APP 대시보드는 사진 안 씀)
          if (hn === '수리전' || hn === '수리후' || hn === '완료확인서' || hn === '확인서' ||
              hn === '수리전_data' || hn === '수리후_data' || hn === '완료확인서_data' || hn === '확인서_data') {
            skipCols[h] = true;
          }
        }
        for (var r = 1; r < vals.length; r++) {
          for (var c = 0; c < vals[r].length; c++) {
            if (skipCols[c]) {
              vals[r][c] = '';
              continue;
            }
            var v = vals[r][c];
            if (v instanceof Date) {
              vals[r][c] = Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
            }
          }
        }
        return {name: ws.getName(), rows: lr, cols: lc, values: vals};
      }

      if (sheetName === '*') {
        var all = ss.getSheets();
        var sheets = {};
        for (var s = 0; s < all.length; s++) {
          var nm = all[s].getName();
          sheets[nm] = gridOf(all[s]);
        }
        return makeRes({status:'ok', sheets: sheets, count: all.length});
      } else {
        var ws = ss.getSheetByName(sheetName);
        if (!ws) return makeRes({status:'error', message:'시트 없음: '+sheetName});
        return makeRes({status:'ok', sheet: gridOf(ws)});
      }
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.20] 일매출 대시보드 PDF 자동 생성 ===
  // ?action=generateDailySalesPdf            (오늘 기준)
  // ?action=generateDailySalesPdf&date=2026-05-05  (특정 날짜)
  if (action === 'generateDailySalesPdf') {
    var dateParam = e.parameter.date || '';
    try {
      var result = generateDailySalesPdf(dateParam);
      return makeRes(Object.assign({status:'ok'}, result));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.20] 시간 트리거 등록 (매일 21:30 KST 자동 실행) — 일회성 ===
  if (action === 'setupDailySalesPdfTrigger') {
    try {
      return makeRes(Object.assign({status:'ok'}, setupDailySalesPdfTrigger()));
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.21] 대시보드 동기화만 호출 (PDF 생성 없이) ===
  // ?action=syncDashboardBeforePdf            (오늘 기준)
  // ?action=syncDashboardBeforePdf&date=2026-05-07
  // [v23.23] 텔레그램 알림 테스트 액션 — Properties에 토큰·chatId 등록 후 통신 확인용
  // ?action=testTelegram&text=hello (text 생략 시 기본 메시지)
  if (action === 'testTelegram') {
    try {
      var props = PropertiesService.getScriptProperties();
      var token = props.getProperty('TG_TOKEN');
      var chatId = props.getProperty('TG_CHAT_ID');
      if (!token) return makeRes({status:'error', message:'TG_TOKEN 미설정 (Script Properties)'});
      if (!chatId) return makeRes({status:'error', message:'TG_CHAT_ID 미설정 (Script Properties)'});
      var text = e.parameter.text || ('🟢 메트로 알림 테스트 — ' +
        Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') + ' KST');
      var resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({chat_id: chatId, text: text, disable_web_page_preview: true}),
        muteHttpExceptions: true
      });
      return makeRes({status:'ok', httpCode: resp.getResponseCode(), body: resp.getContentText()});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  if (action === 'syncDashboardBeforePdf') {
    var dateParam = e.parameter.date || '';
    try {
      return makeRes(Object.assign({status:'ok'}, syncDashboardBeforePdf(dateParam)));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.24] 대시보드 financial 수식 + 월별 매출 추이 차트(2026만) 1회 셋업 ===
  // ?action=setupDashboardFormulas
  if (action === 'setupDashboardFormulas') {
    try {
      return makeRes(Object.assign({status:'ok'}, setupDashboardFormulas()));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.25] 일매출 SUMPRODUCT 범위 확장 (K/L/M/O 끝행 → 2000) ===
  // ?action=extendDailySalesRanges
  if (action === 'extendDailySalesRanges') {
    try {
      return makeRes(Object.assign({status:'ok'}, extendDailySalesRanges()));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.31] 메트로 관리자 일일 보고 메일 (수동 호출) ===
  // ?action=sendDailyReportToManager        (어제 기준)
  // ?action=sendDailyReportToManager&date=2026-05-08  (특정 날짜)
  if (action === 'sendDailyReportToManager') {
    try {
      var dateParam = e.parameter.date || '';
      return makeRes(Object.assign({status:'ok'}, sendDailyReportToManager(dateParam)));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.31] 매일 09:00 KST 메트로 관리자 메일 트리거 등록 ===
  if (action === 'setupManagerReportTrigger') {
    try {
      return makeRes(Object.assign({status:'ok'}, setupManagerReportTrigger()));
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // === [v23.34] 시트 A열 NO 빈 셀 자동 채우기 ===
  // ?action=fillNoSequence&sheetName=파주6단지
  // ?action=fillNoSequence&sheetName=파주6단지&dryRun=1   (미리보기)
  if (action === 'fillNoSequence') {
    try {
      var sn = e.parameter.sheetName || '';
      var dry = (e.parameter.dryRun === '1' || e.parameter.dryRun === 'true');
      if (!sn) return makeRes({status:'error', message:'sheetName 필요'});
      return makeRes(Object.assign({status:'ok'}, fillNoSequence(sn, dry)));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.35] 결제현황 F열 자동화 진단 ===
  // ?action=inspectPaymentSheet
  // 결제현황 F4~F23 (작업비 총액) 수식/값 + 일매출 한 행의 SUMPRODUCT 산식 + 14현장 시트 헤더
  if (action === 'inspectPaymentSheet') {
    try {
      return makeRes(Object.assign({status:'ok'}, inspectPaymentSheet()));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.35] 결제현황 F열 자동집계 수식 일괄 적용 ===
  // ?action=setupPaymentFormulas               (실 적용)
  // ?action=setupPaymentFormulas&dryRun=1      (미리보기, F열 비변경)
  if (action === 'setupPaymentFormulas') {
    try {
      var dry = (e.parameter.dryRun === '1' || e.parameter.dryRun === 'true');
      return makeRes(Object.assign({status:'ok'}, setupPaymentFormulas(dry)));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  // === [v23.35] 사이트 시트 합계 셀 SUM 범위 무한 확장 (진짜 fix) ===
  // ?action=fixSiteTotalRanges               (실 적용)
  // ?action=fixSiteTotalRanges&dryRun=1      (미리보기)
  // 각 사이트 시트의 합계 셀(V21=SUM(V2:V20) 등)을 자기 자신 제외 전체 컬럼 합으로 교체.
  // 합계 행 아래로 새 작업이 추가돼도 자동 누계 — 일매출 SUMPRODUCT v23.25 :$X$2000 확장과 동일 철학.
  if (action === 'fixSiteTotalRanges') {
    try {
      var dry = (e.parameter.dryRun === '1' || e.parameter.dryRun === 'true');
      return makeRes(Object.assign({status:'ok'}, fixSiteTotalRanges(dry)));
    } catch(err) {
      return makeRes({status:'error', message:err.message, stack:err.stack || ''});
    }
  }

  return makeRes({status:'ok', message:'메트로 R&S v23.35 연결됨'});
}

// === [v23.20] 일매출 대시보드 PDF 생성 함수 ===
// _LIVE의 '대시보드' 시트를 PDF로 export → 'A.메트로알엔에스(주)/메트로 당일 매출/' 폴더에 저장
// 파일명: '메트로 당일 매출_yyyy-MM-dd.pdf'
// 시간 트리거에서도 호출됨 — 첫 인자가 event 객체일 수 있어 typeof 체크
function generateDailySalesPdf(targetDateStrOrEvent) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ROOT = 'A.메트로알엔에스(주)';
  var FOLDER = '1.메트로 당일 매출 대시보드';
  var FOLDER_ALIASES = ['메트로 당일 매출 대시보드', '메트로 당일 매출', '메트로_매출_자료전송']; // 옛 이름 fallback (v23.32 → v23.33 prefix 전환 안전망 포함)

  var targetDateStr = (typeof targetDateStrOrEvent === 'string') ? targetDateStrOrEvent : '';
  var today = targetDateStr || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

  // [v23.21] PDF 생성 직전 대시보드 시트 동기화 — 기준일 + financial 영역 자동 갱신
  // 동기화 실패해도 PDF는 옛 값으로 생성되도록 try-catch (재해 시 PDF 자체는 보존)
  var syncResult = null;
  try {
    syncResult = syncDashboardBeforePdf(today);
    SpreadsheetApp.flush();
  } catch (syncErr) {
    Logger.log('[generateDailySalesPdf] sync 실패 (계속 진행): ' + syncErr.message);
  }

  // 대시보드 시트 GID 찾기
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('대시보드');
  if (!ws) throw new Error('대시보드 시트 없음');
  var gid = ws.getSheetId();

  // PDF export URL — Google Sheets 자체 export 엔진
  var url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?' +
    'format=pdf' +
    '&gid=' + gid +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&top_margin=0.5' +
    '&bottom_margin=0.5' +
    '&left_margin=0.4' +
    '&right_margin=0.4' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false';

  var resp = UrlFetchApp.fetch(url, {
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
    muteHttpExceptions: false
  });

  var fileName = '메트로 당일 매출_' + today + '.pdf';
  var blob = resp.getBlob().setName(fileName);

  // Drive 폴더 찾기 — NFC 정규화 매칭 + 옛 이름들 fallback + 자동 생성
  var rootIter = DriveApp.getFoldersByName(ROOT);
  if (!rootIter.hasNext()) throw new Error('루트 폴더 없음: ' + ROOT);
  var rootFolder = rootIter.next();
  var dailyFolder = null;
  var acceptNFC = {};
  acceptNFC[FOLDER.normalize('NFC')] = true;
  for (var ai = 0; ai < FOLDER_ALIASES.length; ai++) {
    acceptNFC[FOLDER_ALIASES[ai].normalize('NFC')] = true;
  }
  var subs = rootFolder.getFolders();
  while (subs.hasNext()) {
    var sub = subs.next();
    var nmNFC = String(sub.getName()).normalize('NFC');
    if (acceptNFC[nmNFC]) {
      dailyFolder = sub;
      break;
    }
  }
  if (!dailyFolder) {
    dailyFolder = rootFolder.createFolder(FOLDER);
  }

  // 같은 이름 기존 PDF 휴지통 (덮어쓰기 효과)
  var existing = dailyFolder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var file = dailyFolder.createFile(blob);

  // [v23.23] 텔레그램 알림 — Script Properties 미설정이면 자동 skip (PDF는 이미 생성됨)
  var tgResult = sendTelegramDailyReport(file, syncResult, today);

  return {
    fileId: file.getId(), fileName: fileName, date: today,
    sizeKB: Math.round(blob.getBytes().length / 1024),
    sync: syncResult,
    telegram: tgResult
  };
}

// === [v23.23] 텔레그램 일일 보고 알림 ===
// generateDailySalesPdf 직후 호출. PropertiesService에 TG_TOKEN + TG_CHAT_ID 설정 필요.
// 미설정/실패 시에도 PDF 생성은 영향 없음 (try-catch로 격리)
function sendTelegramDailyReport(file, syncResult, today) {
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('TG_TOKEN');
    var chatId = props.getProperty('TG_CHAT_ID');
    if (!token || !chatId) {
      Logger.log('[sendTelegramDailyReport] TG_TOKEN/TG_CHAT_ID 미설정 — 알림 건너뜀');
      return {sent: false, reason: 'no_token_or_chatid'};
    }

    var sales = (syncResult && typeof syncResult.todaySales === 'number') ? syncResult.todaySales : 0;
    var unpaid = (syncResult && typeof syncResult.todayUnpaid === 'number') ? syncResult.todayUnpaid : 0;
    var depositCount = (syncResult && syncResult.depositCount) || 0;
    var totalWork = (syncResult && syncResult.totalWork) || 0;
    var payRatePct = syncResult && typeof syncResult.payRate === 'number'
      ? Math.round(syncResult.payRate * 1000) / 10 : 0;
    var todayKor = (syncResult && syncResult.baseDateKor) || today;
    var pdfUrl = file.getUrl();

    var msg = '📊 메트로 일일 보고\n';
    msg += '📅 ' + todayKor + '\n\n';
    msg += '💰 당일 매출: ' + sales.toLocaleString('ko-KR') + '원\n';
    msg += '💸 미지급 잔액: ' + unpaid.toLocaleString('ko-KR') + '원\n';
    msg += '🏗 총 작업비: ' + totalWork.toLocaleString('ko-KR') + '원\n';
    msg += '🏦 입금: ' + depositCount + '회 (입금률 ' + payRatePct + '%)\n\n';
    msg += '📄 PDF 보기:\n' + pdfUrl;

    var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: msg,
        disable_web_page_preview: true
      }),
      muteHttpExceptions: true
    });
    var rc = resp.getResponseCode();
    var body = resp.getContentText().slice(0, 300);
    Logger.log('[sendTelegramDailyReport] HTTP ' + rc + ' / ' + body);
    return {sent: rc === 200, httpCode: rc, body: body};
  } catch (e) {
    Logger.log('[sendTelegramDailyReport] 실패: ' + e.message);
    return {sent: false, error: e.message};
  }
}

// === [v23.24] PDF 생성 직전 대시보드 시트 자동 동기화 ===
// - A3 기준일·당일 매출 텍스트 갱신 (PDF export용 — 텍스트라 수식 불가)
// - financial 영역(M12~M17)은 결제현황 자동 참조 수식이 자동 계산하므로 setValue 안 함
//   * setupDashboardFormulas 액션으로 1회 셋업된 수식이 결제현황 변경 즉시 반영
//   * SpreadsheetApp.flush() 후 셀에서 계산된 최종값을 읽어 텔레그램·PDF 결과로 사용
// - targetDateStr: yyyy-MM-dd (없으면 오늘 KST)
function syncDashboardBeforePdf(targetDateStr) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var TZ = 'Asia/Seoul';

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dash  = ss.getSheetByName('대시보드');
  var sales = ss.getSheetByName('일매출');
  if (!dash)  throw new Error('대시보드 시트 없음');
  if (!sales) throw new Error('일매출 시트 없음');

  var todayStr = targetDateStr || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var todayKor = Utilities.formatDate(
    Utilities.parseDate(todayStr, TZ, 'yyyy-MM-dd'), TZ, 'yyyy년 MM월 dd일'
  );

  // 1) 일매출 시트에서 오늘 매출 추출 (A3 헤더 텍스트 표시용)
  var lastRow = sales.getLastRow();
  var rng = sales.getRange(2, 1, lastRow - 1, 19).getValues();
  var todaySales = 0;
  for (var i = rng.length - 1; i >= 0; i--) {
    var dCell = rng[i][0];
    var dStr = (dCell instanceof Date)
      ? Utilities.formatDate(dCell, TZ, 'yyyy-MM-dd')
      : String(dCell || '').slice(0, 10);
    if (dStr === todayStr) {
      var s = rng[i][18];
      if (typeof s === 'number') todaySales = s;
      break;
    }
    if (dStr && dStr < todayStr) break;
  }

  // 2) A3 기준일·당일 매출 텍스트 갱신
  var headerText = '📅 기준일: ' + todayKor;
  if (todaySales > 0) {
    headerText += '   💰 당일 매출: ' + todaySales.toLocaleString('ko-KR') + '원';
  }
  dash.getRange('A3').setValue(headerText);

  // 3) 시트 수식 즉시 평가 후 financial 셀에서 계산된 최종값 읽기
  //    M12~M17은 결제현황 자동 참조 수식 (setupDashboardFormulas로 셋업됨)
  SpreadsheetApp.flush();
  var fin = dash.getRange('M12:M17').getValues();
  var totalWork    = fin[0][0];  // M12 ='결제현황'!D56 (작업비 VAT별도)
  var depositSumIn = fin[1][0];  // M13 ='결제현황'!C54 (입금 VAT포함)
  var depositSumEx = fin[2][0];  // M14 ='결제현황'!D54 (입금 VAT별도)
  var todayUnpaid  = fin[3][0];  // M15 =M12-M14
  var payRate      = fin[4][0];  // M16 =IFERROR(M14/M12,0)
  var depositCntStr = fin[5][0]; // M17 =COUNT('결제현황'!A29:A1000)&"회"
  var depositCount = parseInt(String(depositCntStr || '').replace(/[^\d]/g, '') || '0', 10);

  return {
    baseDate: todayStr,
    baseDateKor: todayKor,
    todayUnpaid: todayUnpaid,
    todaySales: todaySales,
    totalWork: totalWork,
    depositCount: depositCount,
    depositSumIn: depositSumIn,
    depositSumEx: depositSumEx,
    payRate: typeof payRate === 'number' ? Math.round(payRate * 10000) / 10000 : payRate
  };
}

// === [v23.24] 대시보드 1회성 셋업: financial 수식 + 월별 매출 추이 차트(2026만) ===
// 결제현황 = 단일 진본 패러다임 적용. 한 번 호출하면:
//  (1) M12~M17 = 결제현황 자동 참조 수식 6개 박기 (이후 결제현황 수정 즉시 반영)
//  (2) M16에 % 형식 적용
//  (3) 월별 매출 추이 차트(C29:C42 가까운 차트)에서 2025년 시리즈 제거하고 2026년만 남김
// 멱등(idempotent) — 여러 번 호출해도 동일 결과
function setupDashboardFormulas() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dash = ss.getSheetByName('대시보드');
  if (!dash) throw new Error('대시보드 시트 없음');

  // (1) financial 수식 6개 — 결제현황 자동 합계 셀 직접 참조
  dash.getRange('M12').setFormula("='결제현황'!D56");                      // 작업비 (VAT별도)
  dash.getRange('M13').setFormula("='결제현황'!C54");                      // 입금 (VAT포함)
  dash.getRange('M14').setFormula("='결제현황'!D54");                      // 입금 (VAT별도)
  dash.getRange('M15').setFormula("=M12-M14");                            // 미지급 잔액
  dash.getRange('M16').setFormula("=IFERROR(M14/M12,0)");                  // 입금률
  dash.getRange('M17').setFormula("=COUNT('결제현황'!A29:A1000)&\"회\"");  // 입금 횟수

  // (2) M16 퍼센트 형식
  dash.getRange('M16').setNumberFormat('0.0%');

  // (3) 월별 매출 추이 차트에서 2025년 시리즈 제거
  //    데이터: A29:C42 (A=월, B=2025년, C=2026년). 2025 시리즈를 빼려면 카테고리(A)+값(C)만 남기면 됨.
  var charts = dash.getCharts();
  var chartUpdated = false;
  var chartTitle = '';
  for (var i = 0; i < charts.length; i++) {
    var ch = charts[i];
    var ranges = ch.getRanges();
    var hasMonthlyRange = false;
    for (var r = 0; r < ranges.length; r++) {
      var a1 = ranges[r].getA1Notation();
      // A29:C42 또는 그 안에 포함된 월별 매출 추이 영역
      if (/(?:A|B|C)2[89]/.test(a1) || /(?:A|B|C)3[0-9]/.test(a1) || /(?:A|B|C)4[0-2]/.test(a1)) {
        hasMonthlyRange = true; break;
      }
    }
    if (!hasMonthlyRange) continue;

    var newChart = ch.modify()
      .clearRanges()
      .addRange(dash.getRange('A29:A42'))   // 카테고리 (월)
      .addRange(dash.getRange('C29:C42'))   // 값 (2026년만)
      .build();
    dash.updateChart(newChart);
    chartUpdated = true;
    chartTitle = (ch.getOptions().get('title') || '월별 매출 추이') + '';
    break;
  }

  SpreadsheetApp.flush();
  return {
    formulasSet: ['M12','M13','M14','M15','M16','M17'],
    chartUpdated: chartUpdated,
    chartTitle: chartTitle,
    note: chartUpdated
      ? '월별 매출 추이 차트 2025년 시리즈 제거 완료 (2026년만 표시)'
      : '월별 매출 추이 차트를 자동 식별하지 못함 — 차트 편집기에서 직접 2025년 시리즈 삭제 필요'
  };
}

// === [v23.25] 일매출 시트 SUMPRODUCT 범위 일괄 확장 ===
// K/L/M/O 컬럼의 ':$X$NNN' 끝 행수만 → ':$X$2000' 으로 치환.
// 시작 셀 ':$X$2'는 콜론 앞에 콜론 없으므로 매칭 안 됨 — 보호.
// 단가표 Q/U는 컬럼 화이트리스트(KLMO)에 없어 자동 보호.
// 시트별 hardcoded 행수가 다른 환경(동탄 138, 광주중흥 154, 양주 N 등)에서 모두 일괄 처리.
// 멱등 — 이미 :$X$2000인 셀은 동일 결과로 다시 매칭(변경 없음).
function extendDailySalesRanges() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sales = ss.getSheetByName('일매출');
  if (!sales) throw new Error('일매출 시트 없음');

  var lastRow = sales.getLastRow();
  if (lastRow < 2) return {cellsChanged: 0, note: '일매출 시트 비어있음'};

  // E:R = 14현장 컬럼 (E2~R<lastRow>)
  var range = sales.getRange(2, 5, lastRow - 1, 14);
  var formulas = range.getFormulas();
  var cellsChanged = 0;
  var matchesByCol = {K: 0, L: 0, M: 0, O: 0};

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var f = formulas[i][j];
      if (!f) continue;
      var newF = f;
      ['K','L','M','O'].forEach(function(col) {
        var re = new RegExp(':\\$' + col + '\\$\\d+', 'g');
        var matches = newF.match(re);
        if (matches) {
          matchesByCol[col] += matches.length;
          newF = newF.replace(re, ':$' + col + '$2000');
        }
      });
      if (newF !== f) {
        formulas[i][j] = newF;
        cellsChanged++;
      }
    }
  }

  if (cellsChanged > 0) {
    range.setFormulas(formulas);
    SpreadsheetApp.flush();
  }

  return {
    cellsChanged: cellsChanged,
    matchesByColumn: matchesByCol,
    rangeProcessed: range.getA1Notation(),
    note: 'K/L/M/O 컬럼의 :$X$NNN → :$X$2000 (시작 :$X$2 보호 + 단가표 Q/U 자동 보호)'
  };
}

// === [v23.35] 결제현황 F열 자동화 진단 ===
// 결제현황 F4~F23 21행을 분석:
//   - 현재 F열이 수식인지 hardcoded 값인지
//   - 각 행의 현장명/하자내용/수량 (split row 매핑 단서)
//   - 일매출 시트의 한 데이터 행에서 14현장 컬럼 SUMPRODUCT 산식 dump (패턴 참조용)
//   - 14현장 시트의 헤더 + 마지막 행 번호 (자동집계 SUMPRODUCT 작성 기반)
function inspectPaymentSheet() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pay = ss.getSheetByName('결제현황');
  var sales = ss.getSheetByName('일매출');
  if (!pay) throw new Error('결제현황 시트 없음');
  if (!sales) throw new Error('일매출 시트 없음');

  // 1) 결제현황 A4:F23 — NO/현장/하자/수량/단가/작업비
  var payRange = pay.getRange('A4:F23');
  var payValues = payRange.getValues();
  var payFormulas = payRange.getFormulas();
  var paymentRows = [];
  for (var i = 0; i < payValues.length; i++) {
    paymentRows.push({
      sheetRow: i + 4,
      no: payValues[i][0],
      site: payValues[i][1],
      defect: payValues[i][2],
      qty: payValues[i][3],
      unit: payValues[i][4],
      amount_value: payValues[i][5],
      amount_formula: payFormulas[i][5]
    });
  }
  // 합계 행 (24): 작업비 D24 또는 F24
  var totalRow = pay.getRange('A24:F24').getValues()[0];
  var totalRowFormulas = pay.getRange('A24:F24').getFormulas()[0];

  // 2) 일매출 시트 — 헤더 (row 1) + 한 데이터 행 (row 2) 산식
  var salesLastCol = sales.getLastColumn();
  var salesHeaders = sales.getRange(1, 1, 1, salesLastCol).getValues()[0];
  var salesRow2Values = sales.getRange(2, 1, 1, salesLastCol).getValues()[0];
  var salesRow2Formulas = sales.getRange(2, 1, 1, salesLastCol).getFormulas()[0];
  var salesCols = [];
  for (var c = 0; c < salesHeaders.length; c++) {
    salesCols.push({
      col: c + 1,
      colLetter: columnToLetter_(c + 1),
      header: salesHeaders[c],
      row2_value: salesRow2Values[c],
      row2_formula: salesRow2Formulas[c]
    });
  }

  // 3) 14현장 시트 — 헤더 + 마지막 행 + paymentRows의 산식에서 참조한 셀(V21 등)의 formula·value 추적
  var siteNames = ['경산하양','광주중흥','동탄','양산','양주','원주(무실)','원주(혁신)','충주호암','파주1단지','파주6단지','익산제일','검단제일','감일제일','군산미장'];
  var siteSheets = {};
  for (var si = 0; si < siteNames.length; si++) {
    var sn = siteNames[si];
    var sh = ss.getSheetByName(sn);
    if (!sh) { siteSheets[sn] = {error: '시트 없음'}; continue; }
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    siteSheets[sn] = { headers: hdrs, lastRow: lastRow, lastCol: lastCol };
  }

  // 4) paymentRows의 amount_formula에서 참조된 셀(예: ='양주'!V21*1)을 파싱해 그 셀의 formula+value 가져오기
  var refCellTrace = [];
  for (var pi = 0; pi < paymentRows.length; pi++) {
    var pr = paymentRows[pi];
    var f = pr.amount_formula || '';
    // 패턴: ='시트명'!컬럼행 (예: ='양주'!V21*1, ='파주6단지'!V22*1, ='군산미장'!U16*1)
    var m = f.match(/=\s*'([^']+)'!\s*\$?([A-Z]+)\$?(\d+)/);
    if (m) {
      var refSheet = m[1];
      var refColLetter = m[2];
      var refRow = parseInt(m[3]);
      var rsh = ss.getSheetByName(refSheet);
      if (rsh) {
        try {
          var rng = rsh.getRange(refColLetter + refRow);
          refCellTrace.push({
            paymentSheetRow: pr.sheetRow,
            site: pr.site,
            paymentFormula: f,
            refSheet: refSheet,
            refCell: refColLetter + refRow,
            refValue: rng.getValue(),
            refFormula: rng.getFormula()
          });
        } catch(e) {
          refCellTrace.push({paymentSheetRow: pr.sheetRow, site: pr.site, refSheet: refSheet, refCell: refColLetter + refRow, error: e.message});
        }
      } else {
        refCellTrace.push({paymentSheetRow: pr.sheetRow, site: pr.site, refSheet: refSheet, error: '참조 시트 없음'});
      }
    }
  }

  return {
    paymentRows: paymentRows,
    totalRow24: { values: totalRow, formulas: totalRowFormulas },
    salesCols: salesCols,
    siteSheets: siteSheets,
    refCellTrace: refCellTrace,
    note: '결제현황 F열 자동화 진단 — refCellTrace에 각 사이트 합계 셀(V21 등)의 실제 산식 포함. 산식이 hardcoded면 그게 정합성 누수 원인.'
  };
}

// 컬럼 번호 → A1 알파벳
function columnToLetter_(col) {
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

// === [v23.35] 결제현황 F열 자동집계 수식 일괄 적용 ===
// 21행 split 구조를 14현장 단위로 통합:
//   - 한 현장에 split row가 있으면 첫 번째 row에만 사이트 시트 전체 SUMPRODUCT 박고
//   - 나머지 split row의 F열은 0(또는 ""(빈값))으로 명시
//   - 합계 D56 자체는 그대로 SUM(F4:F23) 유지 (이미 그렇게 됨)
//   - dryRun=true 시 변경 안 하고 계획만 반환
function setupPaymentFormulas(dryRun) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pay = ss.getSheetByName('결제현황');
  var sales = ss.getSheetByName('일매출');
  if (!pay) throw new Error('결제현황 시트 없음');
  if (!sales) throw new Error('일매출 시트 없음');

  // 일매출 시트의 row 2 데이터 행에서 14현장 컬럼의 SUMPRODUCT 산식 추출
  var salesHeaders = sales.getRange(1, 1, 1, sales.getLastColumn()).getValues()[0];
  var salesRow2Formulas = sales.getRange(2, 1, 1, sales.getLastColumn()).getFormulas()[0];

  // 헤더명 → 산식 매핑 (14현장)
  var siteFormulaMap = {};
  for (var c = 0; c < salesHeaders.length; c++) {
    var hn = String(salesHeaders[c] || '').replace(/\s/g, '');
    var f = salesRow2Formulas[c];
    if (f && hn) siteFormulaMap[hn] = f;
  }

  // 결제현황 A4:B23 읽어 현장명별 첫 행 찾기
  var payRange = pay.getRange('A4:F23');
  var payValues = payRange.getValues();
  var payFormulas = payRange.getFormulas();

  // 같은 현장의 첫 row만 자동 SUMPRODUCT, 같은 현장의 둘째/셋째 row는 빈 값
  var seenSites = {};
  var plan = [];
  for (var i = 0; i < payValues.length; i++) {
    var sheetRow = i + 4;
    var site = String(payValues[i][1] || '').trim();
    if (!site) continue;
    var siteKey = site.replace(/\s/g, '');

    // 일매출 산식에서 시트 참조를 ROW 무관 형태로 변환
    var rawFormula = siteFormulaMap[siteKey] || null;
    if (!rawFormula) {
      plan.push({sheetRow: sheetRow, site: site, action: 'skip', reason: '일매출 산식 없음', oldFormula: payFormulas[i][5], oldValue: payValues[i][5]});
      continue;
    }

    if (!seenSites[siteKey]) {
      // 일매출 산식의 ROW($A2), $A2 등 행 참조를 제거 — '연도/월 일치' 조건 빼고 전체 합산
      // 일매출 row2 SUMPRODUCT는 보통: =SUMPRODUCT(('시트'!K:K=ROW관련조건)*('시트'!L:L=조건)*'시트'!M:M)
      // 우리는 전체 누계가 필요하므로 SUMPRODUCT 그대로 두면 일치 0건 → 0. 따라서 시트 SUM 방식으로 새로 작성.
      // 안전한 대안: =SUM('시트명'!M2:M)  ← M이 금액 컬럼이라 가정. 진단 결과 봐야 확정.
      // 일단은 진단 단계라 newFormula은 일매출 산식 그대로 — 사용자 검토 단계에서 결정
      plan.push({
        sheetRow: sheetRow,
        site: site,
        action: 'set-primary',
        oldFormula: payFormulas[i][5],
        oldValue: payValues[i][5],
        newFormula_proposed: rawFormula,
        siteKey: siteKey
      });
      seenSites[siteKey] = true;
    } else {
      plan.push({
        sheetRow: sheetRow,
        site: site,
        action: 'clear-secondary',
        oldFormula: payFormulas[i][5],
        oldValue: payValues[i][5],
        newValue_proposed: 0,
        siteKey: siteKey
      });
    }
  }

  // dryRun이면 적용 안 함
  if (dryRun) {
    return {
      dryRun: true,
      cellsToChange: plan.filter(function(p){return p.action !== 'skip';}).length,
      plan: plan,
      note: '실 적용 전 plan 검토 — newFormula_proposed가 시트 전체 누계가 되는지 확인 필요. 일매출 산식이 ROW 조건 포함이면 다른 산식으로 교체해야 함.'
    };
  }

  // 실 적용은 plan 검토 후 다음 단계에서 별도 함수로 진행
  return {
    dryRun: false,
    note: 'plan 검토 단계 — 실 적용은 별도 진행. 현재는 dryRun=1로만 호출 권장.',
    plan: plan
  };
}

// === [v23.35] 각 사이트 시트의 합계 셀 SUM 범위 무한 확장 ===
// 진단 결과: 양주!V21=`=SUM(V2:V20)`, 파주6단지!V22=`=SUM(V2:V21)` 등 14현장 모두
// 합계 행 위쪽까지만 합산하는 hardcoded 범위. 합계 행 아래로 새 작업이 들어가면 누락.
// 자기 자신 제외 + 1999행까지 무한 확장으로 영구 해결.
//   양주!V21:  =SUM(V2:V20)              → =SUM($V$2:V20)+SUM($V$22:$V$1999)
//   파주6!V22: =SUM(V2:V21)              → =SUM($V$2:V21)+SUM($V$23:$V$1999)
//   익산!V14:  =SUBTOTAL(9,V2:V13)       → =SUBTOTAL(9,$V$2:V13)+SUBTOTAL(9,$V$15:$V$1999)
//   군산!U16:  =SUM(U2:U15)              → =SUM($U$2:U15)+SUM($U$17:$U$1999)
// 멱등 — 이미 확장된 패턴은 다시 변경하지 않음.
function fixSiteTotalRanges(dryRun) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pay = ss.getSheetByName('결제현황');
  if (!pay) throw new Error('결제현황 시트 없음');

  var payFormulas = pay.getRange('F4:F23').getFormulas();
  var results = [];

  for (var i = 0; i < payFormulas.length; i++) {
    var f = payFormulas[i][0] || '';
    var m = f.match(/=\s*'([^']+)'!\s*\$?([A-Z]+)\$?(\d+)/);
    if (!m) continue;

    var refSheet = m[1];
    var refColLetter = m[2];
    var refRow = parseInt(m[3]);
    var rsh = ss.getSheetByName(refSheet);
    if (!rsh) {
      results.push({site: refSheet, status: 'skip', reason: '시트 없음'});
      continue;
    }

    var cell = rsh.getRange(refColLetter + refRow);
    var cur = cell.getFormula();
    if (!cur) {
      results.push({site: refSheet, cell: refColLetter + refRow, status: 'skip', reason: '합계 셀이 hardcoded 값'});
      continue;
    }

    // 이미 fix된 패턴 감지 (1999 들어있으면 skip)
    if (/1999\)/.test(cur)) {
      results.push({site: refSheet, cell: refColLetter + refRow, status: 'already-fixed', current: cur});
      continue;
    }

    var sumMatch = cur.match(/^=\s*SUM\s*\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$/i);
    var subMatch = cur.match(/^=\s*SUBTOTAL\s*\(\s*9\s*,\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$/i);

    var newFormula = null;
    var kind = '';
    if (sumMatch) {
      kind = 'SUM';
      var sc = sumMatch[1].toUpperCase(), sr = parseInt(sumMatch[2]);
      newFormula = '=SUM($' + sc + '$' + sr + ':' + sc + (refRow - 1) + ')+SUM($' + sc + '$' + (refRow + 1) + ':$' + sc + '$1999)';
    } else if (subMatch) {
      kind = 'SUBTOTAL';
      var sc2 = subMatch[1].toUpperCase(), sr2 = parseInt(subMatch[2]);
      newFormula = '=SUBTOTAL(9,$' + sc2 + '$' + sr2 + ':' + sc2 + (refRow - 1) + ')+SUBTOTAL(9,$' + sc2 + '$' + (refRow + 1) + ':$' + sc2 + '$1999)';
    } else {
      results.push({site: refSheet, cell: refColLetter + refRow, status: 'skip', reason: '미인식 패턴', current: cur});
      continue;
    }

    // 예상 값 = 자기 자신 0으로 가정한 전체 컬럼 합 (검증용)
    var preview = null;
    try {
      var fullRange = rsh.getRange(refColLetter + '1:' + refColLetter);
      var colValues = fullRange.getValues();
      var totalNum = 0;
      var curRow = refRow;
      for (var rr = 0; rr < colValues.length; rr++) {
        if (rr + 1 === curRow) continue;
        var vv = colValues[rr][0];
        if (typeof vv === 'number') totalNum += vv;
      }
      preview = totalNum;
    } catch(e) {}

    var oldVal = cell.getValue();
    results.push({
      site: refSheet,
      cell: refColLetter + refRow,
      kind: kind,
      oldFormula: cur,
      oldValue: oldVal,
      newFormula: newFormula,
      previewSum: preview,
      delta: (preview != null && typeof oldVal === 'number') ? (preview - oldVal) : null,
      status: dryRun ? 'planned' : 'applied'
    });

    if (!dryRun) {
      cell.setFormula(newFormula);
    }
  }

  if (!dryRun) SpreadsheetApp.flush();

  return {
    dryRun: dryRun,
    count: results.length,
    appliedCount: results.filter(function(r){return r.status === 'applied';}).length,
    plannedCount: results.filter(function(r){return r.status === 'planned';}).length,
    alreadyFixedCount: results.filter(function(r){return r.status === 'already-fixed';}).length,
    skippedCount: results.filter(function(r){return r.status === 'skip';}).length,
    results: results
  };
}

// === [v23.34] 시트 A열 NO 빈 셀 자동 채우기 ===
// 사용 예: fillNoSequence('파주6단지')
// - 마지막 숫자 NO 찾기 (A열 위에서 아래로)
// - 그 아래 행 중 A는 비었고 B~G 중 하나라도 데이터 있는 행을 대상으로 lastNo+1부터 연속 입력
// - 멱등: 이미 NO 채워진 행은 건너뜀
function fillNoSequence(sheetName, dryRun) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('시트 없음: ' + sheetName);

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {filled: 0, note: '빈 시트'};

  // A:G 한 번에 읽기
  var range = sh.getRange(1, 1, lastRow, 7);
  var values = range.getValues();

  // 마지막 숫자 NO 찾기
  var lastNo = 0;
  var lastNoRow = 0;
  for (var i = 0; i < values.length; i++) {
    var a = values[i][0];
    if (typeof a === 'number' && !isNaN(a)) {
      if (a > lastNo) { lastNo = a; lastNoRow = i + 1; }
    } else if (typeof a === 'string' && a !== '' && a !== 'NO') {
      var n = parseInt(a, 10);
      if (!isNaN(n) && n > lastNo) { lastNo = n; lastNoRow = i + 1; }
    }
  }

  if (lastNo === 0) return {filled: 0, note: 'A열에 숫자 NO 없음'};

  // 채울 대상 행 수집
  var targets = []; // {row, no}
  var nextNo = lastNo + 1;
  for (var j = 0; j < values.length; j++) {
    var row1 = j + 1;
    if (row1 <= lastNoRow) continue; // 마지막 NO 아래만
    var a = values[j][0];
    var hasA = (a !== '' && a !== null);
    if (hasA) continue; // 이미 채워진 행은 보존(멱등)
    // B~G 중 하나라도 데이터 있는 행만 대상
    var hasData = false;
    for (var c = 1; c <= 6; c++) {
      var v = values[j][c];
      if (v !== '' && v !== null) { hasData = true; break; }
    }
    if (!hasData) continue;
    targets.push({row: row1, no: nextNo});
    nextNo++;
  }

  if (targets.length === 0) {
    return {
      filled: 0,
      lastNo: lastNo,
      lastNoRow: lastNoRow,
      note: '채울 빈 NO 행 없음 (모두 정상)'
    };
  }

  // 연속 범위 묶기 (대부분 한 덩어리)
  var firstRow = targets[0].row;
  var lastTargetRow = targets[targets.length - 1].row;
  var contiguous = (lastTargetRow - firstRow + 1 === targets.length);

  var preview = {
    sheetName: sheetName,
    lastNo: lastNo,
    lastNoRow: lastNoRow,
    firstFillRow: firstRow,
    lastFillRow: lastTargetRow,
    firstFillNo: targets[0].no,
    lastFillNo: targets[targets.length - 1].no,
    filled: targets.length,
    contiguous: contiguous
  };

  if (dryRun) {
    preview.dryRun = true;
    return preview;
  }

  if (contiguous) {
    var seq = targets.map(function(t) { return [t.no]; });
    sh.getRange(firstRow, 1, seq.length, 1).setValues(seq);
  } else {
    // 비연속이면 개별 setValue
    for (var k = 0; k < targets.length; k++) {
      sh.getRange(targets[k].row, 1).setValue(targets[k].no);
    }
  }
  SpreadsheetApp.flush();

  return preview;
}

// === [v23.20] 매일 21:30 KST 트리거 등록 (일회성) ===
function setupDailySalesPdfTrigger() {
  // 기존 generateDailySalesPdf 트리거 정리 (중복 방지)
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'generateDailySalesPdf') {
      ScriptApp.deleteTrigger(triggers[t]);
      removed++;
    }
  }
  ScriptApp.newTrigger('generateDailySalesPdf')
    .timeBased()
    .atHour(21)
    .nearMinute(30)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();
  return {message: '매일 21:30 KST generateDailySalesPdf 트리거 등록 완료', removedOld: removed};
}

// === POST 요청 ===
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || '';

    // === 엑셀 데이터 → 구글시트 업로드 ===
    if (action === 'upload') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '하자리스트';
      var headers = body.headers || [];
      var rows = body.rows || [];
      if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});
      if (!headers.length || !rows.length) return makeRes({status:'error', message:'데이터 없음'});

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = ss.getSheetByName(sheetName);
      if (!ws) ws = ss.insertSheet(sheetName);
      else ws.clear();

      ws.getRange(1, 1, 1, headers.length).setValues([headers]);
      ws.getRange(1, 1, 1, headers.length).setFontWeight('bold');

      if (rows.length > 0) {
        var normalizedRows = [];
        for (var i = 0; i < rows.length; i++) {
          var row = rows[i];
          var newRow = [];
          for (var j = 0; j < headers.length; j++) {
            newRow.push(j < row.length ? row[j] : '');
          }
          normalizedRows.push(newRow);
        }
        ws.getRange(2, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
      }

      SpreadsheetApp.flush();
      return makeRes({status:'ok', count:rows.length, sheetName:sheetName});
    }

    // === 사진 저장: Drive 업로드 + =IMAGE() 수식 + _data 열(URL, v23.30) ===
    if (action === 'savePhoto') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      var rowNum = parseInt(body.rowNum);
      var photoType = body.photoType;
      var base64 = body.base64;

      if (!sheetId || !rowNum || !photoType || !base64) {
        return makeRes({status:'error', message:'필수 파라미터 누락'});
      }

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음'});

      var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
      // 시트 헤더가 '완료확인서' 또는 '확인서' 둘 다 지원 (v20.3)
      var typeAliases = {before:['수리전'], after:['수리후'], confirm:['완료확인서','확인서']};
      var aliases = typeAliases[photoType];
      if (!aliases) return makeRes({status:'error', message:'잘못된 photoType: '+photoType});

      var colName = '';
      var imgColIdx = -1;
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        for (var ai = 0; ai < aliases.length; ai++) {
          if (hn === aliases[ai]) { imgColIdx = h + 1; colName = aliases[ai]; break; }
        }
        if (imgColIdx > 0) break;
      }
      // v23.2 안전망: 사진 컬럼 없으면 H~J에 자동 삽입 + 끝에 _data 3개 (oneTimeAddPhotoColumnsToAllSites와 동일 양식)
      if (imgColIdx < 0) {
        ws.insertColumnsBefore(8, 3);
        ws.getRange(1, 8).setValue('수리전').setFontWeight('bold');
        ws.getRange(1, 9).setValue('수리후').setFontWeight('bold');
        ws.getRange(1, 10).setValue('완료확인서').setFontWeight('bold');
        try { ws.setColumnWidth(8, 160); ws.setColumnWidth(9, 160); ws.setColumnWidth(10, 160); } catch(e) {}
        var ec = ws.getLastColumn();
        ws.getRange(1, ec + 1).setValue('수리전_data').setFontWeight('bold');
        ws.getRange(1, ec + 2).setValue('수리후_data').setFontWeight('bold');
        ws.getRange(1, ec + 3).setValue('완료확인서_data').setFontWeight('bold');
        try { ws.hideColumns(ec + 1, 3); } catch(e) {}
        // headers 다시 읽고 imgColIdx 재탐색
        headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
        for (var hh = 0; hh < headers.length; hh++) {
          var hnh = String(headers[hh]).replace(/\s/g,'');
          for (var aii = 0; aii < aliases.length; aii++) {
            if (hnh === aliases[aii]) { imgColIdx = hh + 1; colName = aliases[aii]; break; }
          }
          if (imgColIdx > 0) break;
        }
        if (imgColIdx < 0) return makeRes({status:'error', message:'사진 컬럼 자동 생성 실패 — 시트 양식 점검 필요'});
      }

      var dataColName = colName + '_data';
      var dataColIdx = -1;
      for (var h3 = 0; h3 < headers.length; h3++) {
        if (String(headers[h3]).replace(/\s/g,'') === dataColName) { dataColIdx = h3 + 1; break; }
      }

      // _data 열 자동 생성
      if (dataColIdx < 0) {
        var insertAt = imgColIdx + 1;
        ws.insertColumnAfter(imgColIdx);
        ws.getRange(1, insertAt).setValue(dataColName);
        ws.getRange(1, insertAt).setFontWeight('bold');
        dataColIdx = insertAt;
        ws.hideColumns(dataColIdx);
      }

      // 기존 Drive 파일 있으면 휴지통으로 이동
      tryTrashOldImageFile(ws, rowNum, imgColIdx);

      // v23.1: 행에서 동·호 추출 → 사진 폴더 분기 (A.메트로알엔에스(주)/{현장}/{동}-{호}/)
      var dongVal = '', hoVal = '';
      for (var hd = 0; hd < headers.length; hd++) {
        var hnd = String(headers[hd]).replace(/\s/g,'');
        if (hnd === '동' && !dongVal) dongVal = String(ws.getRange(rowNum, hd+1).getValue() || '').trim();
        else if ((hnd === '호' || hnd === '호수') && !hoVal) hoVal = String(ws.getRange(rowNum, hd+1).getValue() || '').trim();
      }

      // Drive 업로드 + 공개 공유
      var uploaded = uploadPhotoToDrive(base64, sheetId, ws.getName(), rowNum, photoType, dongVal, hoVal);

      // 1) _data 열: Drive URL (v23.30 — 옛 코드는 base64. 셀당 50000자 한계 초과 회피)
      //    read 액션이 'http' 시작 매칭으로 URL/base64 둘 다 처리하므로 호환됨
      ws.getRange(rowNum, dataColIdx).setValue(uploaded.url);

      // 2) 이미지 열: =IMAGE("Drive URL") 수식 (시트 시각용)
      ws.getRange(rowNum, imgColIdx).setFormula('=IMAGE("' + uploaded.url + '")');

      // 이미지 보이게 행 높이 + 컬럼 너비 조정 (v21.10: 가로·세로 모두 160px)
      try { ws.setRowHeight(rowNum, 160); } catch(e) {}
      try { ws.setColumnWidth(imgColIdx, 160); } catch(e) {}

      // v20.6: 확인서 사진 업로드 시 자동 완료 처리
      // v23.6: 첫 매칭(가장 좌측) 우선 — 우측 통계 표의 '완료' 헤더가 잘못 잡히는 버그 수정
      // v23.8: L(완료)은 '완료'가 아닌 모든 값(미완료/N/빈값)을 '완료'로 덮어쓰기 (토글 컬럼이라 보호 불필요)
      //         K(완료일)은 기존 보호 유지 (사용자가 직접 채운 날짜를 덮지 않음)
      var autoDone = false;
      if (photoType === 'confirm') {
        var doneDateIdx = -1, doneIdx = -1;
        for (var hd = 0; hd < headers.length; hd++) {
          var hnd = String(headers[hd]).replace(/\s/g,'');
          if (hnd === '완료일' && doneDateIdx < 0) doneDateIdx = hd + 1;
          else if (hnd === '완료' && doneIdx < 0) doneIdx = hd + 1;
        }
        if (doneDateIdx > 0) {
          var existDate = ws.getRange(rowNum, doneDateIdx).getValue();
          if (!existDate) {
            // v20.7: GAS 프로젝트 시간대가 다를 수 있으므로 명시적 KST 변환
            var ymd = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
            ws.getRange(rowNum, doneDateIdx).setValue(ymd);
            autoDone = true;
          }
        }
        if (doneIdx > 0) {
          var existDone = String(ws.getRange(rowNum, doneIdx).getValue() || '').trim();
          if (existDone !== '완료') {
            ws.getRange(rowNum, doneIdx).setValue('완료');
            autoDone = true;
          }
        }
      }

      SpreadsheetApp.flush();
      return makeRes({status:'ok', url:uploaded.url, fileId:uploaded.id, rowNum:rowNum, photoType:photoType, autoDone:autoDone});
    }

    // === 기존 데이터 마이그레이션 (base64 → Drive 업로드 + =IMAGE) ===
    if (action === 'migratePhotos') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음'});

      var lastRow = ws.getLastRow();
      // v20.3: 시트 헤더가 '완료확인서'/'확인서' 둘 다 지원 — 실제 헤더에서 사용 중인 이름만 처리
      var lastColM = ws.getLastColumn();
      var headersM = ws.getRange(1, 1, 1, lastColM).getValues()[0];
      var headerSetM = {};
      for (var hm = 0; hm < headersM.length; hm++) headerSetM[String(headersM[hm]).replace(/\s/g,'')] = true;
      var typeMap2 = {};
      if (headerSetM['수리전']) typeMap2['수리전'] = 'before';
      if (headerSetM['수리후']) typeMap2['수리후'] = 'after';
      if (headerSetM['완료확인서']) typeMap2['완료확인서'] = 'confirm';
      else if (headerSetM['확인서']) typeMap2['확인서'] = 'confirm';
      var migrated = 0;
      var errors = 0;

      for (var colName2 in typeMap2) {
        var photoType2 = typeMap2[colName2];
        var dataColName2 = colName2 + '_data';
        var imgColIdx2 = -1;
        var dataColIdx2 = -1;

        var lastCol2 = ws.getLastColumn();
        var headers2 = ws.getRange(1, 1, 1, lastCol2).getValues()[0];

        for (var hh = 0; hh < headers2.length; hh++) {
          var hn2 = String(headers2[hh]).replace(/\s/g,'');
          if (hn2 === colName2) imgColIdx2 = hh + 1;
          if (hn2 === dataColName2) dataColIdx2 = hh + 1;
        }
        if (imgColIdx2 < 0) continue;

        if (dataColIdx2 < 0) {
          ws.insertColumnAfter(imgColIdx2);
          ws.getRange(1, imgColIdx2 + 1).setValue(dataColName2);
          ws.getRange(1, imgColIdx2 + 1).setFontWeight('bold');
          dataColIdx2 = imgColIdx2 + 1;
          ws.hideColumns(dataColIdx2);
        }

        for (var r = 2; r <= lastRow; r++) {
          var imgFormula = ws.getRange(r, imgColIdx2).getFormula();
          // 이미 =IMAGE() 수식이면 스킵
          if (imgFormula && imgFormula.indexOf('IMAGE(') >= 0) continue;

          var dataVal = String(ws.getRange(r, dataColIdx2).getValue() || '');
          var imgVal = String(ws.getRange(r, imgColIdx2).getValue() || '');
          var base64v = '';
          if (dataVal.indexOf('data:image') === 0) base64v = dataVal;
          else if (imgVal.indexOf('data:image') === 0) base64v = imgVal;

          if (base64v) {
            try {
              var up = uploadPhotoToDrive(base64v, sheetId, ws.getName(), r, photoType2);
              // v23.30: _data 열에 base64 대신 URL 저장 (셀당 50000자 한계 회피)
              ws.getRange(r, dataColIdx2).setValue(up.url);
              ws.getRange(r, imgColIdx2).setFormula('=IMAGE("' + up.url + '")');
              try { ws.setRowHeight(r, 160); } catch(e) {}  // v21.10: 80 → 160
              try { ws.setColumnWidth(imgColIdx2, 160); } catch(e) {}  // v21.10: 가로 폭도 160
              migrated++;
            } catch(err2) {
              errors++;
            }
          }
        }
      }

      SpreadsheetApp.flush();
      return makeRes({status:'ok', migrated:migrated, errors:errors});
    }

    // === 새 하자 행 추가 (v20.0) — 동시성 보호: LockService ===
    if (action === 'appendRow') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      var data = body.data || {};
      var worker = body.worker || '';

      if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});
      if (!data.dong || !data.ho) return makeRes({status:'error', message:'동/호 필수'});
      if (!data.memo) return makeRes({status:'error', message:'하자내용 필수'});

      // v20.5: standalone Web App에서는 getScriptLock 사용 (getDocumentLock은 container-bound 전용)
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(10000); // 최대 10초 대기 (동시 추가 충돌 방지)
      } catch(le) {
        return makeRes({status:'error', message:'다른 작업자가 추가 중. 잠시 후 다시 시도하세요'});
      }

      try {
        var ss = SpreadsheetApp.openById(sheetId);
        var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
        if (!ws) return makeRes({status:'error', message:'시트 없음: '+sheetName});

        var lastCol = ws.getLastColumn();
        var lastRow = ws.getLastRow();
        var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];

        // 헤더 → 컬럼 인덱스 매핑 (1-based)
        function findCol() {
          var names = Array.prototype.slice.call(arguments);
          for (var i = 0; i < headers.length; i++) {
            var hn = String(headers[i]).replace(/\s/g,'');
            for (var j = 0; j < names.length; j++) {
              if (hn === names[j].replace(/\s/g,'')) return i + 1;
            }
          }
          // 부분 일치 (정확 일치 후순위)
          for (var i2 = 0; i2 < headers.length; i2++) {
            var hn2 = String(headers[i2]).replace(/\s/g,'');
            for (var j2 = 0; j2 < names.length; j2++) {
              if (hn2.indexOf(names[j2].replace(/\s/g,'')) >= 0) return i2 + 1;
            }
          }
          return -1;
        }

        var colNo    = findCol('NO','번호','순번');
        var colDong  = findCol('동');
        var colHo    = findCol('호수','호');
        var colLoc   = findCol('위치','실명');
        var colType  = findCol('유형','하자유형');
        var colMemo  = findCol('하자내용','내용','상세내용');
        var colDate  = findCol('순번','접수일','날짜');
        var colAdded = findCol('작업자','등록자','입력자');

        // v23.10/11: NO 헤더 못 찾으면 A열 fallback. v23.11에선 lastRow가 빈 행이라도
        // 거꾸로 올라가 숫자가 있는 행을 찾아 A열을 NO로 판정 (체인 깨짐 방지)
        var srcRowForFmt = -1;
        if (colNo < 0 && lastRow >= 2) {
          for (var fc = lastRow; fc >= 2; fc--) {
            var v = String(ws.getRange(fc, 1).getValue() || '').trim();
            if (/^\d+$/.test(v)) { colNo = 1; srcRowForFmt = fc; break; }
          }
        } else if (lastRow >= 2) {
          // 헤더 매칭은 됐지만 서식 복사용으로 정상 행 따로 탐색 (마지막 행이 빈 양식일 수 있음)
          for (var fc2 = lastRow; fc2 >= 2; fc2--) {
            var v2 = String(ws.getRange(fc2, colNo > 0 ? colNo : 1).getValue() || '').trim();
            if (v2) { srcRowForFmt = fc2; break; }
          }
        }

        if (colDong < 0 || colHo < 0 || colMemo < 0) {
          return makeRes({status:'error', message:'필수 컬럼 누락 (동/호/하자내용)'});
        }

        // NO 자동 채번 (기존 NO 컬럼 max + 1)
        var nextNo = 1;
        if (colNo > 0 && lastRow >= 2) {
          var nos = ws.getRange(2, colNo, lastRow - 1, 1).getValues();
          for (var n = 0; n < nos.length; n++) {
            var v = parseInt(nos[n][0], 10);
            if (!isNaN(v) && v >= nextNo) nextNo = v + 1;
          }
        }

        // 새 행 작성
        var newRow = new Array(lastCol).fill('');
        if (colNo > 0) newRow[colNo - 1] = nextNo;
        newRow[colDong - 1] = String(data.dong).trim();
        newRow[colHo - 1] = String(data.ho).trim();
        if (colLoc > 0 && data.loc) newRow[colLoc - 1] = String(data.loc).trim();
        if (colType > 0 && data.type) newRow[colType - 1] = String(data.type).trim();
        newRow[colMemo - 1] = String(data.memo).trim();
        if (colDate > 0) {
          // v20.7: 명시적 KST 변환 — GAS 프로젝트 timezone 의존 제거
          newRow[colDate - 1] = Utilities.formatDate(new Date(), 'Asia/Seoul', "M'월'd'일'");
        }
        if (colAdded > 0 && worker) newRow[colAdded - 1] = worker;

        var insertRow = lastRow + 1;
        ws.getRange(insertRow, 1, 1, lastCol).setValues([newRow]);

        // v23.9/11: 정상 행의 서식(테두리·정렬·폰트·배경) 복사
        // v23.11: lastRow가 빈 양식이면 거꾸로 스캔해 찾은 srcRowForFmt 사용 (양식 깨짐 체인 방지)
        var fmtSrc = srcRowForFmt > 0 ? srcRowForFmt : (lastRow >= 2 ? lastRow : -1);
        if (fmtSrc >= 2) {
          try {
            ws.getRange(fmtSrc, 1, 1, lastCol).copyTo(
              ws.getRange(insertRow, 1, 1, lastCol),
              SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
              false
            );
          } catch(fe) { /* 서식 복사 실패해도 값은 들어갔으니 무시 */ }
        }

        SpreadsheetApp.flush();

        return makeRes({
          status:'ok',
          rowNum: insertRow,
          no: nextNo,
          worker: worker
        });
      } finally {
        try { lock.releaseLock(); } catch(re) {}
      }
    }

    // === 사진 삭제 (v20.8) — Drive 파일 trash + 이미지 셀 + _data 셀 비우기 ===
    if (action === 'deletePhoto') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      var rowNum = parseInt(body.rowNum);
      var photoType = body.photoType;

      if (!sheetId || !rowNum || !photoType) {
        return makeRes({status:'error', message:'필수 파라미터 누락'});
      }

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음'});

      var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
      var typeAliases = {before:['수리전'], after:['수리후'], confirm:['완료확인서','확인서']};
      var aliases = typeAliases[photoType];
      if (!aliases) return makeRes({status:'error', message:'잘못된 photoType: '+photoType});

      // 이미지 컬럼 + _data 컬럼 인덱스 찾기 (모든 alias 매칭)
      var imgColIdx = -1, dataColIdx = -1, colName = '';
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        for (var ai = 0; ai < aliases.length; ai++) {
          if (hn === aliases[ai]) { imgColIdx = h + 1; colName = aliases[ai]; break; }
          if (hn === aliases[ai] + '_data') { dataColIdx = h + 1; }
        }
      }
      // _data는 이미지 컬럼명을 기반으로 한 번 더 확인
      if (imgColIdx > 0 && dataColIdx < 0) {
        for (var h2 = 0; h2 < headers.length; h2++) {
          if (String(headers[h2]).replace(/\s/g,'') === colName + '_data') { dataColIdx = h2 + 1; break; }
        }
      }

      // Drive 파일 trash
      if (imgColIdx > 0) tryTrashOldImageFile(ws, rowNum, imgColIdx);

      // 이미지 셀 비우기 (수식 먼저 제거 후 값 비우기)
      if (imgColIdx > 0) {
        ws.getRange(rowNum, imgColIdx).setFormula('');
        ws.getRange(rowNum, imgColIdx).setValue('');
      }

      // _data 셀 비우기
      if (dataColIdx > 0) {
        ws.getRange(rowNum, dataColIdx).setValue('');
      }

      SpreadsheetApp.flush();
      return makeRes({status:'ok', rowNum:rowNum, photoType:photoType, imgCleared:imgColIdx>0, dataCleared:dataColIdx>0});
    }

    // === 행 삭제 (v20.8) — 사진 모두 trash + 행 통째 제거 (LockService 보호) ===
    if (action === 'deleteRow') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      var rowNum = parseInt(body.rowNum);

      if (!sheetId || !rowNum || rowNum < 2) {
        return makeRes({status:'error', message:'rowNum 필수 (2 이상)'});
      }

      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(10000);
      } catch(le) {
        return makeRes({status:'error', message:'다른 작업자가 작업 중. 잠시 후 다시 시도하세요'});
      }

      try {
        var ss = SpreadsheetApp.openById(sheetId);
        var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
        if (!ws) return makeRes({status:'error', message:'시트 없음: '+sheetName});

        var lastRow = ws.getLastRow();
        if (rowNum > lastRow) return makeRes({status:'error', message:'행 번호가 데이터 범위 초과'});

        // 모든 사진 컬럼의 Drive 파일 trash
        var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
        var imgCols = [];
        for (var h = 0; h < headers.length; h++) {
          var hn = String(headers[h]).replace(/\s/g,'');
          if (hn === '수리전' || hn === '수리후' || hn === '완료확인서' || hn === '확인서') {
            imgCols.push(h + 1);
          }
        }
        for (var ic = 0; ic < imgCols.length; ic++) {
          tryTrashOldImageFile(ws, rowNum, imgCols[ic]);
        }

        // 행 삭제
        ws.deleteRow(rowNum);
        SpreadsheetApp.flush();

        return makeRes({status:'ok', rowNum:rowNum, photosTrashed:imgCols.length});
      } finally {
        try { lock.releaseLock(); } catch(re) {}
      }
    }

    return makeRes({status:'error', message:'unknown action: '+action});
  } catch(err) {
    return makeRes({status:'error', message:err.message});
  }
}

// ========================================================================
// === [v23.31] 메트로 관리자 일일 보고 메일 (매일 09:00 KST) ===
// ========================================================================
// 어제 작업분 사진 (Drive 보기 링크) + LIVE 시트 xlsx 사본 (일매출 제외)을
// Script Properties MANAGER_EMAIL 주소로 자동 발송.
// xlsx는 'A.메트로알엔에스(주)/메트로 관리자전송/'에 일별 보관(감사 추적).
// 받는 사람은 다운받아 자유 편집 가능 — 본사 원본 시트엔 영향 없음.
//
// 실행:
//   - GAS 편집기에서 sendDailyReportToManager() 직접 실행 (테스트)
//   - HTTP: ?action=sendDailyReportToManager (어제 기준)
//          ?action=sendDailyReportToManager&date=2026-05-08 (특정일)
//   - 트리거: ?action=setupManagerReportTrigger (1회 등록)
//
// Script Properties 설정 (필수):
//   GAS 편집기 → ⚙ 프로젝트 설정 → 스크립트 속성 → 추가
//     키: MANAGER_EMAIL
//     값: era999@naver.com
function sendDailyReportToManager(targetDateStr) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var TZ = 'Asia/Seoul';

  var props = PropertiesService.getScriptProperties();
  var mgrEmail = props.getProperty('MANAGER_EMAIL');
  if (!mgrEmail) throw new Error('MANAGER_EMAIL Script Property 미설정 — GAS 편집기 ⚙ 프로젝트 설정 → 스크립트 속성에 등록 (예: era999@naver.com)');

  // 트리거 호출 시 첫 인자가 event 객체일 수 있음 — 문자열만 채택
  var explicitDate = (typeof targetDateStr === 'string' && targetDateStr) ? targetDateStr : '';

  // 어제 KST 날짜
  var now = new Date();
  var yStr;
  if (explicitDate) {
    yStr = explicitDate;
  } else {
    var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    yStr = Utilities.formatDate(yesterday, TZ, 'yyyy-MM-dd');
  }
  var yDate = Utilities.parseDate(yStr, TZ, 'yyyy-MM-dd');
  var yDateKor = Utilities.formatDate(yDate, TZ, 'yyyy년 MM월 dd일');
  var dayNames = ['일','월','화','수','목','금','토'];
  var yDayOfWeek = dayNames[yDate.getDay()];
  var todayDateKor = Utilities.formatDate(now, TZ, 'yyyy년 MM월 dd일');

  // 1) 어제 작업 사진 수집 (Drive 폴더 createdDate 기준)
  var photos = collectPhotosForDate(yStr);

  // 2) [v23.32] 어제 사진 통합 폴더 생성 (현장/동호 계층 사본)
  var dailyPhotoFolder = createDailyPhotoArchiveFolder(photos, yStr);
  var dailyPhotoFolderUrl = dailyPhotoFolder ? dailyPhotoFolder.getUrl() : '';

  // 3) LIVE xlsx 사본 생성 (일매출 제외) → 보관 폴더에 저장
  var xlsxFile = exportLiveSheetExcluding(SHEET_ID, ['일매출'], yStr);

  // 4) 메일 본문 (HTML)
  var html = buildDailyReportEmail(photos, yStr, yDateKor, yDayOfWeek, todayDateKor, dailyPhotoFolderUrl);

  // 5) 메일 발송
  MailApp.sendEmail({
    to: mgrEmail,
    subject: '[메트로 R&S] ' + yDateKor + '(' + yDayOfWeek + ') 작업 보고',
    htmlBody: html,
    attachments: [xlsxFile.getBlob()],
    name: '청개구리샤시 본사'
  });

  return {
    sentTo: mgrEmail,
    yesterdayDate: yStr,
    photoCount: photos.totalCount,
    siteCount: photos.sites.length,
    dailyPhotoFolderUrl: dailyPhotoFolderUrl,
    archiveFileId: xlsxFile.getId(),
    archiveFolderUrl: xlsxFile.getParents().next().getUrl()
  };
}

// 어제 사진 수집: A.메트로알엔에스(주)/{현장}/{동-호}/ 폴더 내 createdDate가 어제인 파일
// 시스템 폴더(PDF·관리자 전송 등) 자동 제외
function collectPhotosForDate(targetYmd) {
  var TZ = 'Asia/Seoul';
  var roots = DriveApp.getFoldersByName(DRIVE_PHOTO_ROOT);
  if (!roots.hasNext()) return {sites: [], totalCount: 0};
  var root = roots.next();

  // 작업 사진이 아닌 시스템 폴더 (현장 시트 폴더만 스캔)
  // [v23.33] 새 이름(1. prefix) + 옛 이름 모두 등록 — Drive UI 이름 변경 전·후 모두 안전
  // 'A전체(현장관리)' / '1.A전체(현장관리)'도 포함 (claude현장관리(종합)_LIVE 진본 보호)
  var SKIP = {
    // 새 이름 (v23.33+)
    '1.메트로 당일 매출 대시보드': true,
    '1.메트로 관리자전송': true,
    '1.A전체(현장관리)': true,
    // 옛 이름 (v23.32 이하 — Drive UI 이름 변경 전 안전망)
    '메트로 당일 매출 대시보드': true,
    '메트로_매출_자료전송': true,
    '메트로 당일 매출': true,
    '메트로 관리자전송': true,
    'A전체(현장관리)': true
  };

  var sites = [];
  var totalCount = 0;

  var siteFolders = root.getFolders();
  while (siteFolders.hasNext()) {
    var siteFolder = siteFolders.next();
    var siteName = siteFolder.getName();
    if (SKIP[siteName]) continue;

    var unitFolders = siteFolder.getFolders();
    var siteUnits = [];
    while (unitFolders.hasNext()) {
      var unitFolder = unitFolders.next();
      var unitName = unitFolder.getName(); // '4111-302' 같은 동-호

      var unitPhotos = {before: null, after: null, confirm: null};
      var hasYesterday = false;

      var files = unitFolder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var createdYmd = Utilities.formatDate(file.getDateCreated(), TZ, 'yyyy-MM-dd');
        if (createdYmd !== targetYmd) continue;

        hasYesterday = true;
        var fname = file.getName();
        // savePhoto 명명 규칙: row{N}_{type}_{ts}.{ext}
        if (fname.indexOf('_before_') >= 0) unitPhotos.before = file;
        else if (fname.indexOf('_after_') >= 0) unitPhotos.after = file;
        else if (fname.indexOf('_confirm_') >= 0) unitPhotos.confirm = file;
      }

      if (hasYesterday) {
        var c = 0;
        if (unitPhotos.before) c++;
        if (unitPhotos.after) c++;
        if (unitPhotos.confirm) c++;
        siteUnits.push({
          unit: unitName,
          before: unitPhotos.before,
          after: unitPhotos.after,
          confirm: unitPhotos.confirm,
          count: c
        });
        totalCount += c;
      }
    }

    if (siteUnits.length > 0) {
      sites.push({siteName: siteName, units: siteUnits});
    }
  }

  return {sites: sites, totalCount: totalCount};
}

// LIVE 시트 xlsx 사본 생성 (특정 시트 제외) → 'A.메트로알엔에스(주)/메트로 관리자전송/'에 보관
// 반환: 보관된 xlsx 파일 (Blob 추출 가능)
// idempotent — 같은 날짜 같은 이름 파일은 휴지통 후 새로 생성
function exportLiveSheetExcluding(sheetId, excludeSheetNames, ymd) {
  // [v23.33] 새 이름 우선, 옛 이름 fallback (Drive UI 이름 변경 전·후 모두 정상 동작)
  var ARCHIVE = '1.메트로 관리자전송';
  var ARCHIVE_ALIASES = ['메트로 관리자전송'];

  // 보관 폴더 확보 (새 이름 → 옛 이름 → 새 이름으로 자동 생성)
  var roots = DriveApp.getFoldersByName(DRIVE_PHOTO_ROOT);
  if (!roots.hasNext()) throw new Error('루트 폴더 없음: ' + DRIVE_PHOTO_ROOT);
  var rootFolder = roots.next();
  var archiveFolder = _findOrCreateSubFolder(rootFolder, ARCHIVE, ARCHIVE_ALIASES);

  var copyName = 'claude현장관리_' + ymd;

  // 같은 이름 기존 xlsx 휴지통 (중복 방지)
  var existing = archiveFolder.getFilesByName(copyName + '.xlsx');
  while (existing.hasNext()) existing.next().setTrashed(true);

  // 1) 임시 Google Sheets 사본 (LIVE는 그대로 — 이게 핵심: 원본 보호)
  var srcFile = DriveApp.getFileById(sheetId);
  var tmpCopy = srcFile.makeCopy(copyName + '_tmp', archiveFolder);

  try {
    // 2) 사본에서 일매출 등 제외 시트 삭제
    var copySS = SpreadsheetApp.openById(tmpCopy.getId());
    excludeSheetNames.forEach(function(sn) {
      var ws = copySS.getSheetByName(sn);
      if (ws) copySS.deleteSheet(ws);
    });
    SpreadsheetApp.flush();

    // 3) xlsx로 export
    var url = 'https://docs.google.com/spreadsheets/d/' + tmpCopy.getId() + '/export?format=xlsx';
    var resp = UrlFetchApp.fetch(url, {
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}
    });
    var xlsxBlob = resp.getBlob().setName(copyName + '.xlsx');

    // 4) 보관 폴더에 xlsx 저장 (감사 추적용)
    var savedFile = archiveFolder.createFile(xlsxBlob);
    return savedFile;

  } finally {
    // 5) 임시 Google Sheets 사본은 휴지통 (보관 안 함, xlsx만 남김)
    tmpCopy.setTrashed(true);
  }
}

// [v23.32] 어제 사진을 통합 폴더에 사본 복사 → 매니저가 폴더 링크 1번으로 통째 다운로드
// 구조: A.메트로알엔에스(주)/메트로 관리자전송/사진_yyyy-MM-dd/{현장명}/{동-호}/{수리전|수리후|확인서}.{ext}
// 원본 동호 폴더는 보존 (사본만 추가). 같은 날짜 재실행 시 휴지통 후 재생성 (멱등).
// 폴더에 ANYONE_WITH_LINK + VIEW 권한 → 링크 받은 사람 누구나 다운로드 가능 (편집 불가)
// photos가 비어있으면 null 반환 (메일 본문에서 건너뜀)
function createDailyPhotoArchiveFolder(photos, ymd) {
  if (!photos || !photos.sites || photos.sites.length === 0) return null;

  // [v23.33] 새 이름 우선, 옛 이름 fallback (Drive UI 이름 변경 전·후 모두 정상 동작)
  var ARCHIVE = '1.메트로 관리자전송';
  var ARCHIVE_ALIASES = ['메트로 관리자전송'];

  var roots = DriveApp.getFoldersByName(DRIVE_PHOTO_ROOT);
  if (!roots.hasNext()) throw new Error('루트 폴더 없음: ' + DRIVE_PHOTO_ROOT);
  var rootFolder = roots.next();
  var archiveFolder = _findOrCreateSubFolder(rootFolder, ARCHIVE, ARCHIVE_ALIASES);

  var dailyFolderName = '사진_' + ymd;

  // 멱등: 같은 이름 폴더 있으면 휴지통 후 새로 생성
  var existing = archiveFolder.getFoldersByName(dailyFolderName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var dailyFolder = archiveFolder.createFolder(dailyFolderName);

  // 현장 → 동호 → 사진 사본
  photos.sites.forEach(function(site) {
    var siteSubFolder = dailyFolder.createFolder(site.siteName);
    site.units.forEach(function(unit) {
      var unitSubFolder = siteSubFolder.createFolder(unit.unit);
      if (unit.before)  copyWithLabel_(unit.before,  '수리전',  unitSubFolder);
      if (unit.after)   copyWithLabel_(unit.after,   '수리후',  unitSubFolder);
      if (unit.confirm) copyWithLabel_(unit.confirm, '확인서',  unitSubFolder);
    });
  });

  // 폴더 자체에 ANYONE_WITH_LINK + VIEW (Drive는 폴더 권한이 내부 파일에도 상속됨)
  try {
    dailyFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    // 일부 도메인 정책에서 ANYONE_WITH_LINK 차단 가능 — 그래도 폴더는 생성됨
    Logger.log('폴더 공유 권한 설정 실패 (수동 공유 필요): ' + e.message);
  }

  return dailyFolder;
}

// 사본 파일 이름을 '수리전.jpg' 등 의미 있는 이름으로 (확장자 보존)
function copyWithLabel_(srcFile, label, destFolder) {
  var origName = srcFile.getName();
  var dot = origName.lastIndexOf('.');
  var ext = (dot >= 0) ? origName.substring(dot) : '.jpg';
  srcFile.makeCopy(label + ext, destFolder);
}

// HTML 메일 본문 생성
function buildDailyReportEmail(photos, yStr, yDateKor, yDayOfWeek, todayDateKor, dailyPhotoFolderUrl) {
  var html = '<div style="font-family:\'Noto Sans KR\',\'Malgun Gothic\',sans-serif;max-width:700px;color:#333;line-height:1.6;">';
  html += '<p>안녕하세요, 구미영 대리님.</p>';
  html += '<p><b>' + yDateKor + ' (' + yDayOfWeek + ')</b> 메트로 R&S 작업 보고 송부드립니다.</p>';

  if (photos.sites.length === 0) {
    html += '<p style="color:#999;background:#f5f5f5;padding:10px;border-radius:6px;">📷 어제 신규 등록된 작업 사진이 없습니다.</p>';
  } else {
    // [v23.32] 사진 통합 폴더 링크 — 메일 최상단에 강조 박스
    if (dailyPhotoFolderUrl) {
      html += '<div style="background:#FFFBEB;border:2px solid #F59E0B;border-radius:8px;padding:14px 18px;margin:18px 0;">';
      html += '<p style="margin:0 0 8px 0;font-size:15px;"><b>📁 어제 작업 사진 통합 폴더 (현장별 정리)</b></p>';
      html += '<p style="margin:0 0 10px 0;color:#666;font-size:13px;">아래 링크를 클릭하시면 현장 → 동·호 폴더가 정리된 화면이 열립니다. 우측 상단 <b>다운로드</b> 버튼을 누르면 전체를 한번에 zip 파일로 받으실 수 있습니다.</p>';
      html += '<p style="margin:0;"><a href="' + dailyPhotoFolderUrl + '" style="display:inline-block;background:#F59E0B;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold;">📥 사진 폴더 열기 / 다운로드</a></p>';
      html += '</div>';
    }

    // 작업 요약
    html += '<h3 style="color:#0F172A;border-bottom:2px solid #0891B2;padding-bottom:5px;margin-top:25px;">▶ 작업 요약</h3>';
    html += '<ul>';
    var totalUnits = 0;
    photos.sites.forEach(function(s) {
      html += '<li><b>' + s.siteName + '</b>: ' + s.units.length + '건</li>';
      totalUnits += s.units.length;
    });
    html += '</ul>';
    html += '<p style="background:#ECFEFF;padding:10px;border-left:3px solid #0891B2;margin:10px 0;"><b>총 ' + totalUnits + '건 (사진 ' + photos.totalCount + '장)</b></p>';

    // 사진 링크 — 현장별 그룹
    html += '<h3 style="color:#0F172A;border-bottom:2px solid #0891B2;padding-bottom:5px;margin-top:25px;">▶ 사진 링크 (클릭하면 큰 사진)</h3>';
    photos.sites.forEach(function(s) {
      html += '<h4 style="margin-top:18px;color:#334155;background:#F1F5F9;padding:6px 10px;border-radius:4px;">📍 ' + s.siteName + '</h4>';
      s.units.forEach(function(u) {
        html += '<p style="margin:6px 0 4px 20px;"><b>' + u.unit + '</b> &nbsp;';
        if (u.before) html += '<a href="' + u.before.getUrl() + '" style="color:#0891B2;text-decoration:none;margin-right:8px;">[수리전]</a>';
        if (u.after) html += '<a href="' + u.after.getUrl() + '" style="color:#0891B2;text-decoration:none;margin-right:8px;">[수리후]</a>';
        if (u.confirm) html += '<a href="' + u.confirm.getUrl() + '" style="color:#0891B2;text-decoration:none;">[확인서]</a>';
        html += '</p>';
      });
    });
  }

  // 첨부 안내
  html += '<h3 style="color:#0F172A;border-bottom:2px solid #0891B2;padding-bottom:5px;margin-top:25px;">▶ 시트 첨부</h3>';
  html += '<p><b>📎 claude현장관리_' + yStr + '.xlsx</b> <span style="color:#666;">(일매출 시트 제외)</span></p>';
  html += '<p style="color:#666;font-size:13px;background:#FFF8E1;padding:10px;border-radius:4px;border-left:3px solid #F59E0B;margin-top:8px;">';
  html += '※ 첨부 xlsx 파일은 어제 시점의 <b>스냅샷 사본</b>입니다. 다운로드 후 자유롭게 편집하셔도 본사 원본 시트에는 영향이 없습니다.';
  html += '</p>';

  // 푸터
  html += '<hr style="margin-top:35px;border:none;border-top:1px solid #ddd;">';
  html += '<p style="color:#666;font-size:12px;line-height:1.6;">';
  html += todayDateKor + ' 09:00 자동 발송<br>';
  html += '청개구리샤시 본사 (frogsash.co.kr)<br>';
  html += '경기도 의왕시 시청로 42, 108동 1702호';
  html += '</p>';
  html += '</div>';
  return html;
}

// 매일 09:00 KST 트리거 등록 (1회성 — 멱등)
function setupManagerReportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'sendDailyReportToManager') {
      ScriptApp.deleteTrigger(triggers[t]);
      removed++;
    }
  }
  ScriptApp.newTrigger('sendDailyReportToManager')
    .timeBased()
    .atHour(9)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();
  return {
    message: '매일 09:00 KST sendDailyReportToManager 트리거 등록 완료',
    removedOld: removed
  };
}
