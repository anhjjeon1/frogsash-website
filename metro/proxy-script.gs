// ========================================
// (주)메트로 R&S AI v16.0 - Google Apps Script
// 구글시트 협업 + 사진 시트 직접 저장 (DriveApp 미사용)
// ========================================

function makeRes(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// === GET 요청 ===
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  // 구글시트 읽기 (사진 base64 포함)
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

        // 사진 셀 값 직접 읽기 (base64 data URL)
        obj._photos = {};
        for (var pType in photoCols) {
          var col = photoCols[pType];
          var val = String(data[i][col] || '');
          if (val.indexOf('data:image') === 0) {
            obj._photos[pType] = val;
          }
        }
        rows.push(obj);
      }
      return makeRes({status:'ok', rows:rows, sheetName:ws.getName(), count:rows.length});
    } catch(err) {
      return makeRes({status:'error', message:err.message});
    }
  }

  return makeRes({status:'ok', message:'메트로 R&S v16.0 연결됨'});
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

    // === 사진 → 시트 셀에 직접 저장 (base64) ===
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

      var colIdx = -1;
      for (var h = 0; h < headers.length; h++) {
        if (String(headers[h]).replace(/\s/g,'') === colName) { colIdx = h + 1; break; }
      }
      if (colIdx < 0) return makeRes({status:'error', message:colName+' 열 없음'});

      // 셀에 base64 직접 저장 (DriveApp 불필요)
      ws.getRange(rowNum, colIdx).setValue(base64);
      SpreadsheetApp.flush();

      return makeRes({status:'ok', url:base64, rowNum:rowNum, photoType:photoType});
    }

    return makeRes({status:'error', message:'unknown action: '+action});
  } catch(err) {
    return makeRes({status:'error', message:err.message});
  }
}
