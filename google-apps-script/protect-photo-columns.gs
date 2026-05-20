/**
 * 일회용 보호 설정 — _LIVE 14현장 시트의 H/I/J(사진 컬럼) 우발 Delete 차단.
 *
 * 사고 경위 (2026-05-20):
 *  - 시트 공유 권한은 소유자(사장님) 단독으로 이미 잠겨 있었음
 *  - 그럼에도 파주6단지·양주·양산 H/I/J IMAGE 수식 77셀 누락 사고 발생
 *  - GAS deletePhoto 코드 결백 (formula+value+_data 모두 함께 비움)
 *  - 결론: 소유자 본인의 시트 우발 작업 (필터+Delete 추정)
 *
 * 안전성:
 *  - setWarningOnly(true) — 경고만, 의도적 편집은 확인 한 번 더 받고 통과
 *  - GAS owner 권한 자동 통과 (savePhoto·deletePhoto·migratePhotos 정상)
 *  - 헤더 행(1행)은 보호 범위에서 제외
 *  - 사진 컬럼이 없는 시스템 시트(대시보드·일매출·결제현황)는 자동 skip
 *  - 멱등 재실행 안전 (기존 동일 description 보호는 갱신)
 *
 * 사용:
 *  1) protectPhotoColumnsAllSheets()  실행
 *  2) 실행 로그에서 시트별 protected 상태 확인
 *  3) 검증: 임의 시트의 H/I/J 셀 클릭 → Delete → 경고 다이얼로그 뜨는지 확인
 */

function protectPhotoColumnsAllSheets() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var PROT_DESC = '🛡️ 자동 사진 컬럼 보호 — H/I/J(수리전·수리후·완료확인서) 우발 Delete 방지. GAS는 owner 권한 통과';

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var report = [];

  for (var s = 0; s < sheets.length; s++) {
    var ws = sheets[s];
    var name = ws.getName();
    var lc = ws.getLastColumn();
    if (lc < 10) { report.push({sheet: name, status: 'skipped (컬럼 부족)'}); continue; }

    var headers = ws.getRange(1, 1, 1, lc).getValues()[0];
    var cols = {};
    var labelAlias = {'확인서': '완료확인서'};
    for (var h = 0; h < headers.length; h++) {
      var hn = String(headers[h]).replace(/\s/g, '');
      if (labelAlias[hn]) hn = labelAlias[hn];
      if (!(hn in cols)) cols[hn] = h + 1;
    }

    if (!cols['수리전'] || !cols['수리후'] || !cols['완료확인서']) {
      report.push({sheet: name, status: 'skipped (사진 컬럼 없음)'});
      continue;
    }

    // 기존 동일 description 보호 제거 (멱등)
    var existing = ws.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var removed = 0;
    for (var e = 0; e < existing.length; e++) {
      if (String(existing[e].getDescription() || '').indexOf('자동 사진 컬럼 보호') >= 0) {
        existing[e].remove();
        removed++;
      }
    }

    // 사진 컬럼 인덱스 (대부분 8/9/10 연속, 일부 시트 다를 수 있음)
    var imgCols = [cols['수리전'], cols['수리후'], cols['완료확인서']].sort(function(a,b){return a-b;});
    var maxRows = ws.getMaxRows();

    // 연속 3컬럼이면 단일 범위, 아니면 분리
    var protectedCount = 0;
    if (imgCols[2] - imgCols[0] === 2) {
      var range = ws.getRange(2, imgCols[0], maxRows - 1, 3);
      var prot = range.protect();
      prot.setDescription(PROT_DESC);
      prot.setWarningOnly(true);
      protectedCount = 1;
    } else {
      for (var i = 0; i < imgCols.length; i++) {
        var range2 = ws.getRange(2, imgCols[i], maxRows - 1, 1);
        var prot2 = range2.protect();
        prot2.setDescription(PROT_DESC);
        prot2.setWarningOnly(true);
        protectedCount++;
      }
    }

    report.push({
      sheet: name,
      status: 'protected',
      imgCols: imgCols,
      ranges: protectedCount,
      removedExisting: removed
    });
  }

  var summary = {
    timestamp: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    totalSheets: sheets.length,
    protectedSheets: report.filter(function(r){return r.status === 'protected';}).length,
    perSheet: report
  };

  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}

/**
 * 보호 해제 (필요 시 — 예: 시트 양식 일괄 수정 작업 전 임시 해제)
 */
function unprotectPhotoColumnsAllSheets() {
  var SHEET_ID = '1xyAXLOINOVpTLhw21qO0I6IHqVzBhQHfutDN4QNa2Q4';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var removed = 0;
  for (var s = 0; s < sheets.length; s++) {
    var prots = sheets[s].getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var p = 0; p < prots.length; p++) {
      if (String(prots[p].getDescription() || '').indexOf('자동 사진 컬럼 보호') >= 0) {
        prots[p].remove();
        removed++;
      }
    }
  }
  Logger.log('removed ' + removed + ' protections');
  return {removed: removed};
}
