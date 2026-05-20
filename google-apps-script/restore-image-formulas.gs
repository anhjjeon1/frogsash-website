/**
 * 일회용 복구 스크립트 — _LIVE 14시트 H/I/J(수리전·수리후·완료확인서)의
 * `=IMAGE(url)` 수식이 누락된 셀을, AD/AE/AF(_data 컬럼)의 URL을 사용해 복원.
 *
 * 발견 경위 (2026-05-20):
 *  - 파주6단지 NO 1258·1259·1276·1283·1285·1286 inspectCell 결과,
 *    _data 컬럼엔 Drive URL이 정상 존재 / H·I·J formula=""
 *  - 사진 자체는 Drive에 안전, 시트 시각 표시만 깨짐 (필터+Delete 의심)
 *
 * 안전성:
 *  - 빈 셀에만 setFormula (기존 수식·값 절대 덮어쓰지 않음)
 *  - _data URL 무수정
 *  - Drive 파일 무접근
 *  - dryRun 모드로 카운트만 미리보기 가능
 *
 * 사용:
 *  1) restoreImagesAllSheets_DRY()  먼저 실행 (변경 없이 카운트)
 *  2) Logger 출력 확인 (보기 > 실행 로그)
 *  3) restoreImagesAllSheets()       실제 복원 적용
 */

function restoreImagesAllSheets_DRY() { return _restoreImagesCore(true); }
function restoreImagesAllSheets()      { return _restoreImagesCore(false); }

function _restoreImagesCore(dryRun) {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var allSheets = ss.getSheets();
  var perSheet = [];
  var totalFound = 0;
  var totalRestored = 0;

  for (var s = 0; s < allSheets.length; s++) {
    var ws = allSheets[s];
    var name = ws.getName();
    var lr = ws.getLastRow();
    var lc = ws.getLastColumn();
    if (lr < 2 || lc < 10) continue;

    var headers = ws.getRange(1, 1, 1, lc).getValues()[0];

    // 컬럼 헤더 매핑 (첫 매칭 우선, '확인서'→'완료확인서' alias)
    var cols = {};
    var labelAlias = {'확인서': '완료확인서', '확인서_data': '완료확인서_data'};
    for (var h = 0; h < headers.length; h++) {
      var hn = String(headers[h]).replace(/\s/g, '');
      if (labelAlias[hn]) hn = labelAlias[hn];
      if (!(hn in cols)) cols[hn] = h + 1;
    }

    var triples = [
      ['수리전',     '수리전_data'],
      ['수리후',     '수리후_data'],
      ['완료확인서', '완료확인서_data']
    ];

    // 사진 양식을 갖춘 시트만 처리 (대시보드·일매출·결제현황 등 시스템 시트 자동 skip)
    var hasFull = true;
    for (var i = 0; i < triples.length; i++) {
      if (!cols[triples[i][0]] || !cols[triples[i][1]]) { hasFull = false; break; }
    }
    if (!hasFull) {
      perSheet.push({sheet: name, status: 'skipped (사진 컬럼 없음)', found: 0, restored: 0});
      continue;
    }

    var sheetFound = 0;
    var sheetRestored = 0;
    var rowsToBump = {};
    var perCol = {};

    for (var t = 0; t < triples.length; t++) {
      var imgCol = cols[triples[t][0]];
      var dataCol = cols[triples[t][1]];

      var imgFormulas = ws.getRange(2, imgCol, lr - 1, 1).getFormulas();
      var imgValues   = ws.getRange(2, imgCol, lr - 1, 1).getValues();
      var dataValues  = ws.getRange(2, dataCol, lr - 1, 1).getValues();

      var colFound = 0;
      var colRestored = 0;

      for (var r = 0; r < dataValues.length; r++) {
        var url = String(dataValues[r][0] || '').trim();
        var formula = String(imgFormulas[r][0] || '').trim();
        var value   = String(imgValues[r][0] || '').trim();
        // 복원 조건: _data에 http URL 존재 AND 이미지 셀이 완전히 비어있음
        if (url.indexOf('http') === 0 && !formula && !value) {
          colFound++;
          sheetFound++;
          if (!dryRun) {
            var rowNum = r + 2;
            ws.getRange(rowNum, imgCol).setFormula('=IMAGE("' + url + '")');
            colRestored++;
            sheetRestored++;
            rowsToBump[rowNum] = true;
          }
        }
      }

      perCol[triples[t][0]] = {found: colFound, restored: colRestored};
    }

    // 복원된 행은 사진 보이게 row height 160 (try-catch로 안전)
    if (!dryRun) {
      for (var rn in rowsToBump) {
        try { ws.setRowHeight(parseInt(rn), 160); } catch(e) {}
      }
    }

    if (sheetFound > 0) {
      perSheet.push({
        sheet: name, lastRow: lr, found: sheetFound, restored: sheetRestored,
        byCol: perCol
      });
      totalFound += sheetFound;
      totalRestored += sheetRestored;
    } else {
      perSheet.push({sheet: name, status: 'clean (복원 대상 없음)', found: 0, restored: 0});
    }
  }

  if (!dryRun) SpreadsheetApp.flush();

  var summary = {
    mode: dryRun ? '🔬 DRY_RUN (변경 없음)' : '✅ APPLIED (실제 복원됨)',
    timestamp: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    totalSheets: allSheets.length,
    sheetsAffected: perSheet.filter(function(p){return p.found > 0;}).length,
    totalCellsToRestore: totalFound,
    totalCellsRestored: totalRestored,
    perSheet: perSheet
  };

  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}
