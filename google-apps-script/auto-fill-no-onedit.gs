/**
 * 시트 자동 NO 채움 — onEdit 단순 트리거
 *
 * 동작:
 *  사용자가 _LIVE 시트의 B/C/D/E/F/G 컬럼(순번/동/호수/위치/유형/하자내용)에 직접 입력 시
 *  같은 행 A열(NO)이 비어있으면 자동으로 lastNo+1 채움.
 *  다중 셀 paste 지원 (행별 NO 증분).
 *
 * 대상 시트:
 *  A1 헤더가 'NO'이고 사진 컬럼(수리전 등)이 있는 14현장 시트만
 *  대시보드·일매출·결제현황·단가표·전체공정표·2025년매출·2026년매출 등 시스템 시트는 자동 skip
 *
 * 트리거:
 *  단순 onEdit — 함수 이름 onEdit() 그대로 두면 자동 활성화 (별도 트리거 등록 불필요)
 *  GAS 편집기에 파일 추가·저장 후 즉시 작동
 *
 * 안전:
 *  - 이미 NO가 있는 행은 무시 (덮어쓰지 않음)
 *  - 헤더 행(1행) 무시
 *  - 사진/_data 컬럼(H~AF 영역) 편집은 무시 — B~G(2~7) 편집 시에만 발화
 *  - 편집된 행의 B~G에 실제 값이 없으면 무시 (Delete만 한 경우 NO 안 채움)
 *  - 무한 루프 방지: 스크립트의 setValue 호출은 onEdit 트리거 안 됨 (GAS 동작)
 *  - 오류 발생 시 silent (단순 트리거 throw 무시됨, console.log만)
 *
 * 한계:
 *  - 단순 트리거 권한 제한 — 외부 API·UrlFetch 불가 (우리 케이스 setValue만이라 무관)
 *  - 30초 실행 한도 (대량 paste도 충분, 145행 paste = ~수 초)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var ws = e.range.getSheet();

    // 시트 양식 검증 — A1='NO' 헤더 + 사진 컬럼 있는 현장 시트만
    var headerA1 = String(ws.getRange(1, 1).getValue() || '').trim();
    if (headerA1 !== 'NO') return;

    var lc = ws.getLastColumn();
    if (lc < 10) return;

    var headers = ws.getRange(1, 1, 1, lc).getValues()[0];
    var hasPhotoCol = false;
    for (var h = 0; h < headers.length; h++) {
      var hn = String(headers[h]).replace(/\s/g, '');
      if (hn === '수리전' || hn === '수리후' || hn === '완료확인서' || hn === '확인서') {
        hasPhotoCol = true;
        break;
      }
    }
    if (!hasPhotoCol) return;

    // 편집 범위
    var startRow = e.range.getRow();
    var startCol = e.range.getColumn();
    var numRows = e.range.getNumRows();
    var numCols = e.range.getNumColumns();
    var endCol = startCol + numCols - 1;

    // 편집 범위가 B~G(2~7) 컬럼과 겹쳐야 발화
    if (endCol < 2 || startCol > 7) return;

    // 마지막 NO 계산 (한 번만)
    var lastRow = ws.getLastRow();
    var maxNo = 0;
    if (lastRow >= 2) {
      var noVals = ws.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var rv = 0; rv < noVals.length; rv++) {
        var v = parseInt(noVals[rv][0]);
        if (!isNaN(v) && v > maxNo) maxNo = v;
      }
    }

    // 편집된 각 행에 대해 A열 자동 채움
    var nextNo = maxNo + 1;
    for (var row = startRow; row < startRow + numRows; row++) {
      if (row < 2) continue; // 헤더 skip

      var noCell = ws.getRange(row, 1);
      var existing = noCell.getValue();
      if (existing !== '' && existing !== null && existing !== undefined) continue;

      // 추가 안전: 편집한 행의 B~G에 실제 값이 있어야 채움 (Delete만 한 경우 무시)
      var rowBG = ws.getRange(row, 2, 1, 6).getValues()[0];
      var hasData = false;
      for (var c = 0; c < rowBG.length; c++) {
        if (String(rowBG[c]).trim() !== '') { hasData = true; break; }
      }
      if (!hasData) continue;

      noCell.setValue(nextNo);
      nextNo++;
    }
  } catch (err) {
    // 단순 트리거 throw 무시됨, 로그만
    console.log('[onEdit auto-NO] ' + (err && err.message ? err.message : err));
  }
}
