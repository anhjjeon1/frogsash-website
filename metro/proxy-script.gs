// ========================================
// (주)메트로 R&S AI v18.0 - Google Apps Script
// 구글시트 협업 + Drive 사진 업로드 + =IMAGE() 수식 표시
// ========================================

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
        else if (hn === '완료확인서_data') photoDataCols.confirm = h;
        else if (hn === '수리전') photoImgCols.before = h;
        else if (hn === '수리후') photoImgCols.after = h;
        else if (hn === '완료확인서') photoImgCols.confirm = h;
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
      var typeMap = {before:'수리전', after:'수리후', confirm:'완료확인서'};
      var colName = typeMap[photoType];
      if (!colName) return makeRes({status:'error', message:'잘못된 photoType: '+photoType});

      var dataColName = colName + '_data';
      var imgColIdx = -1;
      var dataColIdx = -1;
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        if (hn === colName) imgColIdx = h + 1;
        if (hn === dataColName) dataColIdx = h + 1;
      }
      if (imgColIdx < 0) return makeRes({status:'error', message:colName+' 열 없음'});

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

      SpreadsheetApp.flush();
      return makeRes({status:'ok', url:uploaded.url, fileId:uploaded.id, rowNum:rowNum, photoType:photoType});
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
      var typeMap2 = {'수리전':'before','수리후':'after','완료확인서':'confirm'};
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

    return makeRes({status:'error', message:'unknown action: '+action});
  } catch(err) {
    return makeRes({status:'error', message:err.message});
  }
}
