// ========================================
// (주)메트로 R&S AI v18.0 - Google Apps Script
// 구글시트 협업 + Drive 사진 업로드 + =IMAGE() 수식 표시
// ========================================

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

// Drive 내에 현장별 사진 폴더 확보
function getOrCreatePhotoFolder(sheetId, sheetName) {
  var rootName = 'METRO_PHOTOS';
  var roots = DriveApp.getFoldersByName(rootName);
  var root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
  var subName = (sheetName || 'sheet') + '_' + String(sheetId).substring(0, 8);
  var subs = root.getFoldersByName(subName);
  return subs.hasNext() ? subs.next() : root.createFolder(subName);
}

// base64 → Drive 업로드 + 공개 공유
function uploadPhotoToDrive(base64, sheetId, sheetName, rowNum, photoType) {
  var match = base64.match(/data:(.*?);base64,(.*)/);
  if (!match) throw new Error('잘못된 base64 데이터');
  var mime = match[1];
  var bytes = Utilities.base64Decode(match[2]);
  var ext = (mime.split('/')[1] || 'jpg').split('+')[0];
  var fname = 'row' + rowNum + '_' + photoType + '_' + new Date().getTime() + '.' + ext;
  var blob = Utilities.newBlob(bytes, mime, fname);
  var folder = getOrCreatePhotoFolder(sheetId, sheetName);
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return {
    id: file.getId(),
    url: 'https://lh3.googleusercontent.com/d/' + file.getId()
  };
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

  return makeRes({status:'ok', message:'메트로 R&S v18.0 연결됨'});
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
      if (imgColIdx < 0) return makeRes({status:'error', message:aliases.join('/')+' 열 없음'});

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

      // Drive 업로드 + 공개 공유
      var uploaded = uploadPhotoToDrive(base64, sheetId, ws.getName(), rowNum, photoType);

      // 1) _data 열: base64 (앱/엑셀 읽기용)
      ws.getRange(rowNum, dataColIdx).setValue(base64);

      // 2) 이미지 열: =IMAGE("Drive URL") 수식 (시트 시각용)
      ws.getRange(rowNum, imgColIdx).setFormula('=IMAGE("' + uploaded.url + '")');

      // 이미지 보이게 행 높이 조정
      try { ws.setRowHeight(rowNum, 80); } catch(e) {}

      // v20.6: 확인서 사진 업로드 시 자동 완료 처리 (빈 셀일 때만 — 기존 데이터 보호)
      var autoDone = false;
      if (photoType === 'confirm') {
        var doneDateIdx = -1, doneIdx = -1;
        for (var hd = 0; hd < headers.length; hd++) {
          var hnd = String(headers[hd]).replace(/\s/g,'');
          if (hnd === '완료일') doneDateIdx = hd + 1;
          else if (hnd === '완료') doneIdx = hd + 1;
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
          var existDone = ws.getRange(rowNum, doneIdx).getValue();
          if (!existDone) {
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
              try { ws.setRowHeight(r, 80); } catch(e) {}
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
