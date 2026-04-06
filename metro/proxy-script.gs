// ========================================
// (주)메트로 R&S AI v13.0 - Google Apps Script
// 구글시트 협업 + 사진 Drive 즉시 저장
// ========================================

function makeRes(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// === base64 → Google Drive 저장 → 썸네일 URL 반환 ===
function savePhotoToDrive(base64DataUrl, fileName) {
  if (!base64DataUrl || base64DataUrl.indexOf('data:image') < 0) return '';
  try {
    var parts = base64DataUrl.split(',');
    var mime = parts[0].match(/:(.*?);/)[1];
    var bytes = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(bytes, mime, fileName);

    var folders = DriveApp.getFoldersByName('메트로_사진');
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('메트로_사진');

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w200';
  } catch(err) {
    return 'DRIVE_ERR:' + err.message;
  }
}

// === GET 요청 ===
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  // 구글시트 읽기 (사진 URL 포함)
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

      // 사진 컬럼 인덱스 찾기
      var photoCols = {};
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        if (hn === '수리전') photoCols.before = h;
        else if (hn === '수리후') photoCols.after = h;
        else if (hn === '완료확인서') photoCols.confirm = h;
      }

      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          var isPh = false;
          for (var pt in photoCols) { if (photoCols[pt] === j) { isPh = true; break; } }
          if (isPh) {
            obj[headers[j]] = '';
          } else {
            obj[headers[j]] = data[i][j] !== undefined ? String(data[i][j]) : '';
          }
        }
        obj._rowNum = i + 1;

        // IMAGE 수식에서 URL 추출
        obj._photos = {};
        for (var pType in photoCols) {
          var col = photoCols[pType];
          var formula = formulas[i][col] || '';
          if (formula) {
            var match = formula.match(/IMAGE\s*\(\s*"([^"]+)"/i);
            if (match) obj._photos[pType] = match[1];
          }
        }
        rows.push(obj);
      }
      return makeRes({status:'ok', rows:rows, sheetName:ws.getName(), count:rows.length});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  // 파일 프록시 (CORS 우회)
  if (action === 'proxy') {
    var url = e.parameter.url;
    if (!url) return makeRes({status:'error', message:'url 필요'});
    try {
      var resp = UrlFetchApp.fetch(url, {followRedirects:true, muteHttpExceptions:true});
      if (resp.getResponseCode() !== 200) return makeRes({status:'error', message:'HTTP '+resp.getResponseCode()});
      var base64 = Utilities.base64Encode(resp.getBlob().getBytes());
      return makeRes({status:'ok', data:base64});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  return makeRes({status:'ok', message:'메트로 R&S v13.0 연결됨'});
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

      // 헤더 쓰기
      ws.getRange(1, 1, 1, headers.length).setValues([headers]);
      ws.getRange(1, 1, 1, headers.length).setFontWeight('bold');

      // 데이터 쓰기
      if (rows.length > 0) {
        // 각 행의 길이를 헤더와 맞추기
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

    // === 개별 사진 → Drive 저장 + 시트 IMAGE 수식 ===
    if (action === 'savePhoto') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      var rowNum = parseInt(body.rowNum);
      var photoType = body.photoType; // before, after, confirm
      var base64 = body.base64;
      var worker = body.worker || '';

      if (!sheetId || !rowNum || !photoType || !base64) {
        return makeRes({status:'error', message:'필수 파라미터 누락'});
      }

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음'});

      // 사진 컬럼 찾기
      var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
      var typeMap = {before:'수리전', after:'수리후', confirm:'완료확인서'};
      var colName = typeMap[photoType];
      if (!colName) return makeRes({status:'error', message:'잘못된 photoType: '+photoType});

      var colIdx = -1;
      for (var h = 0; h < headers.length; h++) {
        if (String(headers[h]).replace(/\s/g,'') === colName) { colIdx = h + 1; break; }
      }
      if (colIdx < 0) return makeRes({status:'error', message:colName+' 열 없음'});

      // Drive에 사진 저장
      var prefix = (sheetName || '현장') + '_R' + rowNum + '_' + colName;
      if (worker) prefix += '_' + worker;
      var url = savePhotoToDrive(base64, prefix + '.jpg');
      if (!url || url.indexOf('DRIVE_ERR') === 0) return makeRes({status:'error', message:url || 'Drive 저장 실패 (빈 응답)'});

      // 시트에 IMAGE 수식 삽입
      ws.getRange(rowNum, colIdx).setFormula('=IMAGE("' + url + '")');
      SpreadsheetApp.flush();

      return makeRes({status:'ok', url:url, rowNum:rowNum, photoType:photoType});
    }

    // === 배치 동기화 (레거시 호환) ===
    if (action === 'sync') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      var items = body.items || [];
      if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});
      if (!items.length) return makeRes({status:'error', message:'items 비어있음'});

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음'});

      var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
      var colIdx = {};
      for (var h = 0; h < headers.length; h++) {
        var hName = String(headers[h]).replace(/\s/g,'');
        if (hName === '수리전') colIdx.before = h + 1;
        else if (hName === '수리후') colIdx.after = h + 1;
        else if (hName === '완료확인서') colIdx.confirm = h + 1;
        else if (hName === '완료일') colIdx.doneDate = h + 1;
        else if (hName === '완료') colIdx.done = h + 1;
      }

      var updated = 0;
      for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var row = item._rowNum;
        if (!row || row < 2) continue;
        var prefix = (sheetName||'현장') + '_' + row + '_';

        if (colIdx.before && item.before && item.before.indexOf('data:image') >= 0) {
          var urlB = savePhotoToDrive(item.before, prefix + '수리전.jpg');
          if (urlB) ws.getRange(row, colIdx.before).setFormula('=IMAGE("' + urlB + '")');
        }
        if (colIdx.after && item.after && item.after.indexOf('data:image') >= 0) {
          var urlA = savePhotoToDrive(item.after, prefix + '수리후.jpg');
          if (urlA) ws.getRange(row, colIdx.after).setFormula('=IMAGE("' + urlA + '")');
        }
        if (colIdx.confirm && item.confirm && item.confirm.indexOf('data:image') >= 0) {
          var urlC = savePhotoToDrive(item.confirm, prefix + '완료확인서.jpg');
          if (urlC) ws.getRange(row, colIdx.confirm).setFormula('=IMAGE("' + urlC + '")');
        }
        if (colIdx.doneDate && item.date) ws.getRange(row, colIdx.doneDate).setValue(item.date);
        if (colIdx.done) ws.getRange(row, colIdx.done).setValue('완료');
        updated++;
      }
      SpreadsheetApp.flush();
      return makeRes({status:'ok', updated:updated});
    }

    return makeRes({status:'error', message:'unknown action: '+action});
  } catch(err) {
    return makeRes({status:'error', message:err.message});
  }
}
