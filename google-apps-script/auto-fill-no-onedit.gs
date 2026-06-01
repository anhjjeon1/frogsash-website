/**
 * _LIVE onEdit 단순 트리거 — 두 가지 자동화 통합
 *
 *  ① 자동 NO 채움
 *     B/C/D/E/F/G(순번/동/호수/위치/유형/하자내용) 직접 입력 시
 *     같은 행 A열(NO)이 비어있으면 lastNo+1 자동 채움. 다중 paste 지원.
 *
 *  ② 사진 IMAGE 수식 자가복구 (v2 — 2026-06-01 추가)
 *     H/I/J(수리전/수리후/완료확인서)를 시트에서 직접 Delete 해 빈 셀이 됐는데
 *     짝꿍 _data 열(수리전_data 등)에 Drive URL이 살아있으면
 *     즉시 =IMAGE(url) 자동 복원. 필터 작업 중 우발적 H:J Delete로
 *     사진이 시트에서 사라지던 반복 사고(2026-05-20·05-28·06-01)를 영구 차단.
 *     경고창 없음 — 사장님 작업 흐름 방해 0.
 *
 * 대상 시트:
 *  A1 헤더가 'NO'이고 사진 컬럼(수리전 등)이 있는 14현장 시트만
 *  대시보드·일매출·결제현황·단가표·전체공정표·매출 등 시스템 시트는 자동 skip
 *
 * 트리거:
 *  단순 onEdit — 함수 이름 onEdit() 그대로 두면 자동 활성화 (별도 등록 불필요)
 *  ※ 단순 트리거 활성화는 파일 추가·저장 후 최대 5분 지연될 수 있음
 *
 * 안전:
 *  - 스크립트의 setValue/setFormula 호출은 onEdit 재발화 안 함 → 무한 루프 없음
 *  - 앱(GAS web)의 사진 삭제(deletePhoto)는 _data까지 비우므로 자가복구 대상 아님
 *    (복구 조건이 'http URL 존재'라 _data가 비면 복원 안 함 — 정상 삭제 존중)
 *  - 자가복구는 '빈 셀 + _data에 http URL'일 때만 setFormula (기존 수식 안 덮어씀)
 *  - 단순 트리거 권한 제한(UrlFetch 불가)과 무관 — setValue/setFormula만 사용
 *  - 오류 발생 시 silent (단순 트리거 throw 무시됨, console.log만)
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
    var numRows  = e.range.getNumRows();
    var numCols  = e.range.getNumColumns();
    var endCol   = startCol + numCols - 1;

    // ===== ① 사진 IMAGE 수식 자가복구 =====
    // 편집 범위가 사진 이미지 컬럼(H/I/J)과 겹칠 때만 검사
    _selfHealPhotoFormulas(ws, headers, startRow, numRows, startCol, endCol);

    // ===== ② 자동 NO 채움 =====
    // 편집 범위가 B~G(2~7) 컬럼과 겹쳐야 발화
    if (endCol < 2 || startCol > 7) return;

    var lastRow = ws.getLastRow();
    var maxNo = 0;
    if (lastRow >= 2) {
      var noVals = ws.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var rv = 0; rv < noVals.length; rv++) {
        var v = parseInt(noVals[rv][0]);
        if (!isNaN(v) && v > maxNo) maxNo = v;
      }
    }

    var nextNo = maxNo + 1;
    for (var row = startRow; row < startRow + numRows; row++) {
      if (row < 2) continue; // 헤더 skip

      var noCell = ws.getRange(row, 1);
      var existing = noCell.getValue();
      if (existing !== '' && existing !== null && existing !== undefined) continue;

      // 편집한 행의 B~G에 실제 값이 있어야 채움 (Delete만 한 경우 무시)
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
    console.log('[onEdit] ' + (err && err.message ? err.message : err));
  }
}

/**
 * 사진 이미지 셀(H/I/J)이 빈 값이 됐는데 짝꿍 _data에 URL이 남아있으면 =IMAGE(url) 복원.
 * 헤더 이름으로 컬럼을 동적 해석 (시트마다 위치 달라도 안전).
 */
function _selfHealPhotoFormulas(ws, headers, startRow, numRows, startCol, endCol) {
  try {
    // 이미지 컬럼명 → _data 컬럼명 매핑
    var pairs = [
      { img: '수리전',     data: '수리전_data' },
      { img: '수리후',     data: '수리후_data' },
      { img: '완료확인서', data: '완료확인서_data', altImg: '확인서', altData: '확인서_data' }
    ];

    // 헤더 인덱스(1-based) 해석 — 첫 매칭 우선
    var idx = {};
    for (var h = 0; h < headers.length; h++) {
      var hn = String(headers[h]).replace(/\s/g, '');
      if (!(hn in idx)) idx[hn] = h + 1;
    }

    for (var p = 0; p < pairs.length; p++) {
      var imgName  = pairs[p].img;
      var dataName = pairs[p].data;
      var imgCol  = idx[imgName]  || (pairs[p].altImg  ? idx[pairs[p].altImg]  : 0);
      var dataCol = idx[dataName] || (pairs[p].altData ? idx[pairs[p].altData] : 0);
      if (!imgCol || !dataCol) continue;

      // 이 이미지 컬럼이 편집 범위와 겹치지 않으면 skip
      if (imgCol < startCol || imgCol > endCol) continue;

      for (var row = startRow; row < startRow + numRows; row++) {
        if (row < 2) continue;

        var imgCell = ws.getRange(row, imgCol);
        var curFormula = String(imgCell.getFormula() || '').trim();
        var curValue   = String(imgCell.getValue() || '').trim();
        if (curFormula || curValue) continue; // 비어있을 때만 복원

        var url = String(ws.getRange(row, dataCol).getValue() || '').trim();
        if (url.indexOf('http') === 0) {
          imgCell.setFormula('=IMAGE("' + url + '")');
          try { ws.setRowHeight(row, 160); } catch (e2) {}
          console.log('[selfHeal] ' + ws.getName() + ' ' + imgCell.getA1Notation() + ' 복원');
        }
      }
    }
  } catch (err) {
    console.log('[selfHeal] ' + (err && err.message ? err.message : err));
  }
}
