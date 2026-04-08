// ========================================
// (주)메트로 R&S AI v17.0 - Google Apps Script
// 구글시트 협업 + 사진 CellImage 표시 + _data 열 저장
// ========================================

function makeRes(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// === GET 요청 ===
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  // 구글시트 읽기 (사진 base64는 _data 열에서 읽기)
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
      var headers = data[0];

      // 사진 _data 컬럼 인덱스 찾기 (base64 텍스트가 저장된 열)
      var photoDataCols = {};
      // 사진 이미지 컬럼 인덱스 (CellImage가 저장된 열 - 스킵용)
      var photoImgCols = {};
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        // _data 열 (base64 텍스트)
        if (hn === '수리전_data') photoDataCols.before = h;
        else if (hn === '수리후_data') photoDataCols.after = h;
        else if (hn === '완료확인서_data') photoDataCols.confirm = h;
        // 이미지 열 (CellImage - 값 읽기 불필요)
        else if (hn === '수리전') photoImgCols.before = h;
        else if (hn === '수리후') photoImgCols.after = h;
        else if (hn === '완료확인서') photoImgCols.confirm = h;
      }

      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          // 사진 이미지 열과 _data 열은 일반 데이터에서 제외
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

        // _data 열에서 사진 base64 읽기
        obj._photos = {};
        for (var pType in photoDataCols) {
          var col = photoDataCols[pType];
          var val = String(data[i][col] || '');
          if (val.indexOf('data:image') === 0) {
            obj._photos[pType] = val;
          }
        }

        // 하위 호환: _data 열이 없으면 기존 이미지 열에서 읽기 시도
        if (!photoDataCols.before && photoImgCols.before !== undefined) {
          var val = String(data[i][photoImgCols.before] || '');
          if (val.indexOf('data:image') === 0) obj._photos.before = val;
        }
        if (!photoDataCols.after && photoImgCols.after !== undefined) {
          var val = String(data[i][photoImgCols.after] || '');
          if (val.indexOf('data:image') === 0) obj._photos.after = val;
        }
        if (!photoDataCols.confirm && photoImgCols.confirm !== undefined) {
          var val = String(data[i][photoImgCols.confirm] || '');
          if (val.indexOf('data:image') === 0) obj._photos.confirm = val;
        }

        rows.push(obj);
      }
      return makeRes({status:'ok', rows:rows, sheetName:ws.getName(), count:rows.length});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  return makeRes({status:'ok', message:'메트로 R&S v17.0 연결됨'});
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

    // === 사진 저장: CellImage(시각용) + _data 열(base64 텍스트) ===
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

      // 이미지 열 인덱스 찾기
      var imgColIdx = -1;
      var dataColIdx = -1;
      for (var h = 0; h < headers.length; h++) {
        var hn = String(headers[h]).replace(/\s/g,'');
        if (hn === colName) imgColIdx = h + 1;
        if (hn === dataColName) dataColIdx = h + 1;
      }

      if (imgColIdx < 0) return makeRes({status:'error', message:colName+' 열 없음'});

      // _data 열이 없으면 자동 생성 (이미지 열 바로 뒤)
      if (dataColIdx < 0) {
        var insertAt = imgColIdx + 1;
        ws.insertColumnAfter(imgColIdx);
        ws.getRange(1, insertAt).setValue(dataColName);
        ws.getRange(1, insertAt).setFontWeight('bold');
        dataColIdx = insertAt;
        // 열 숨기기
        ws.hideColumns(dataColIdx);
      }

      // 1) _data 열에 base64 텍스트 저장 (앱 읽기용)
      ws.getRange(rowNum, dataColIdx).setValue(base64);

      // 2) 이미지 열에 CellImage 저장 (구글시트 시각용)
      try {
        var cellImage = SpreadsheetApp.newCellImage()
          .setSourceUrl(base64)
          .setAltTextTitle(colName + ' (No.' + rowNum + ')')
          .build();
        ws.getRange(rowNum, imgColIdx).setValue(cellImage);
      } catch(imgErr) {
        // CellImage 실패 시 base64 텍스트라도 저장
        ws.getRange(rowNum, imgColIdx).setValue(base64);
      }

      SpreadsheetApp.flush();
      return makeRes({status:'ok', url:base64, rowNum:rowNum, photoType:photoType});
    }

    // === 기존 데이터 마이그레이션 (base64 텍스트 → CellImage + _data) ===
    if (action === 'migratePhotos') {
      var sheetId = body.sheetId;
      var sheetName = body.sheetName || '';
      if (!sheetId) return makeRes({status:'error', message:'sheetId 필요'});

      var ss = SpreadsheetApp.openById(sheetId);
      var ws = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
      if (!ws) return makeRes({status:'error', message:'시트를 찾을 수 없음'});

      var lastRow = ws.getLastRow();
      var lastCol = ws.getLastColumn();
      var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];

      var photoTypes = ['수리전','수리후','완료확인서'];
      var migrated = 0;

      for (var t = 0; t < photoTypes.length; t++) {
        var colName = photoTypes[t];
        var dataColName = colName + '_data';
        var imgColIdx = -1;
        var dataColIdx = -1;

        // 현재 헤더 다시 읽기 (열 삽입 후 변경될 수 있음)
        lastCol = ws.getLastColumn();
        headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];

        for (var h = 0; h < headers.length; h++) {
          var hn = String(headers[h]).replace(/\s/g,'');
          if (hn === colName) imgColIdx = h + 1;
          if (hn === dataColName) dataColIdx = h + 1;
        }

        if (imgColIdx < 0) continue;

        // _data 열 없으면 생성
        if (dataColIdx < 0) {
          ws.insertColumnAfter(imgColIdx);
          ws.getRange(1, imgColIdx + 1).setValue(dataColName);
          ws.getRange(1, imgColIdx + 1).setFontWeight('bold');
          dataColIdx = imgColIdx + 1;
          ws.hideColumns(dataColIdx);
          // 헤더 다시 읽기
          lastCol = ws.getLastColumn();
          headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
          // imgColIdx는 그대로, dataColIdx 업데이트
        }

        // 기존 base64 텍스트를 CellImage로 변환
        for (var r = 2; r <= lastRow; r++) {
          var val = String(ws.getRange(r, imgColIdx).getValue() || '');
          if (val.indexOf('data:image') === 0) {
            // _data 열에 base64 복사
            ws.getRange(r, dataColIdx).setValue(val);
            // 이미지 열에 CellImage 설정
            try {
              var cellImage = SpreadsheetApp.newCellImage()
                .setSourceUrl(val)
                .setAltTextTitle(colName + ' (No.' + r + ')')
                .build();
              ws.getRange(r, imgColIdx).setValue(cellImage);
              migrated++;
            } catch(imgErr) {
              // CellImage 변환 실패 시 텍스트 유지
            }
          }
        }
      }

      SpreadsheetApp.flush();
      return makeRes({status:'ok', migrated:migrated});
    }

    return makeRes({status:'error', message:'unknown action: '+action});
  } catch(err) {
    return makeRes({status:'error', message:err.message});
  }
}
