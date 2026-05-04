// ========================================
// (주)메트로 R&S AI v23.9 - Google Apps Script
// 구글시트 협업 + Drive 사진 업로드/삭제 + 행 추가/삭제 + =IMAGE() 수식 표시
// 액션: read, upload, savePhoto, migratePhotos, appendRow, deletePhoto, deleteRow, listSheets, checkCompleteColumns
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

        // 1순위: _data 열 (base64 텍스트)
        for (var pType in photoDataCols) {
          var col = photoDataCols[pType];
          var val = String(data[i][col] || '');
          if (val.indexOf('data:image') === 0 || val.indexOf('http') === 0) {
            obj._photos[pType] = val;
          }
        }

        // 2순위: 이미지 열 수식에서 Drive URL 추출 (v18+ =IMAGE())
        for (var pType2 in photoImgCols) {
          if (obj._photos[pType2]) continue;
          var colI = photoImgCols[pType2];
          var fm = (formulas[i] && formulas[i][colI]) ? String(formulas[i][colI]) : '';
          if (fm) {
            var um = fm.match(/"(https?:\/\/[^"]+)"/);
            if (um) { obj._photos[pType2] = um[1]; continue; }
          }
          // 하위 호환: 이미지 열에 base64 텍스트가 그대로 있는 경우
          var raw = String(data[i][colI] || '');
          if (raw.indexOf('data:image') === 0 || raw.indexOf('http') === 0) {
            obj._photos[pType2] = raw;
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

  return makeRes({status:'ok', message:'메트로 R&S v23.9 연결됨'});
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

    // === 사진 저장: Drive 업로드 + =IMAGE() 수식 + _data 열(base64) ===
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

      // 1) _data 열: base64 (앱/엑셀 읽기용)
      ws.getRange(rowNum, dataColIdx).setValue(base64);

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
              ws.getRange(r, dataColIdx2).setValue(base64v);
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

        // v23.9: 이전 행 서식(테두리·정렬·폰트·배경) 새 행에 복사 — 시트 양식 일관성 유지
        if (lastRow >= 2) {
          try {
            ws.getRange(lastRow, 1, 1, lastCol).copyTo(
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
