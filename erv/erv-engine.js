/* 청개구리 환기량 적정성 계산 엔진 (ERV Engine) — 결로진단앱(A)·전용앱(B) 공용
 * 원칙: 판정 숫자는 이 파일의 공식만 만든다(AI는 설명만). 근거 없는 값은 쓰지 않고, 사용자가 입력하지 않은 값은 계산에서 빼고 표시한다.
 * 근거
 *  - 법정 환기량: 「건축물의 설비기준 등에 관한 규칙」 제11조 — 신축 공동주택 시간당 0.5회 이상
 *  - 유효환기량: 같은 규칙 별표 1의5 제15호 — 열회수형 환기장치 유효환기량은 표시용량의 90% 이상
 *  - 1인당 25㎥/h: 2006년 같은 규칙이 다중이용시설에 정한 값(주택 기준 아님, 비교용)
 *  - 수분 발생량: BS 5250:2011 Table D.6 (Typical moisture generation rates for household activities)
 *  - 바깥 공기 습도: Open-Meteo 과거 기상(지난겨울 12~2월 일평균 기온·상대습도)
 */
(function (root) {
  'use strict';
  var ENGINE_VERSION = 'erv-engine 1.0';

  // ── 습공기 (Magnus 식, 결로진단앱과 동일)
  function es(T) { return 0.6108 * Math.exp(17.27 * T / (T + 237.3)); }          // 포화수증기압 kPa
  function mixRatio(T, RH) { var e = es(T) * RH / 100; return 622 * e / (101.325 - e); } // 절대습도 g/kg(건공기)
  function allowRH(Ti, Ts) { return Math.min(100, es(Ts) / es(Ti) * 100); }        // 표면온도 Ts에서 결로가 시작되는 실내 습도
  var RHO = 1.2; // 공기 밀도 kg/㎥ (20℃ 부근)

  // ── BS 5250:2011 Table D.6 값
  var BS = {
    asleep: 40,        // g/h·인
    seated: 70,        // g/h·인 (앉아서 활동)
    cookGas: 3000,     // g/일
    cookElec: 2000,    // g/일
    dish: 400,         // g/일
    shower: 600,       // g/회 (15분)
    washClothes: 500,  // g/일
    dryIndoor: 1500,   // g/일
    floors: 200,       // g/일
    plant: 20          // g/일·개
  };

  function num(v) { var n = parseFloat(v); return isNaN(n) ? null : n; }

  /* 입력(in)
   *  area ㎡, height m, people 명, homeHours(평일 집에 있는 시간, 수면 포함) h, sleepHours h
   *  cook: 'gas'|'elec'|'none', cookDays(주 n일), showersPerDay, laundryDays(주 n일 세탁), dry: 'indoor'|'hp'|'exhaust'|'balc'|''
   *  plants 개, humidMlh ml/h, humidHours h, aquarium bool, newborn bool, cantOpen bool(맞벌이 등 창문 환기 어려움)
   *  window: 'passive'|'system'|'double'|'old'|'' , Ti 실내온도, targetRH 목표 상한 습도, Ts 측정 유리 표면온도(선택)
   *  outT, outRH (겨울 바깥 평균), product: {low,mid,high CMH, type:'erv'|'hrv'|'', latentEff %, filter:'ok'|'due'|'clog'|'', error bool}
   */
  function compute(inp) {
    var r = { version: ENGINE_VERSION, missing: [], notes: [], sources: [] };
    var A = num(inp.area), H = num(inp.height), N = num(inp.people);
    if (!A || !H) r.missing.push('면적·천장 높이');
    if (!N) r.missing.push('거주 인원');
    r.volume = (A && H) ? A * H : null;

    // 1) 법정
    r.qLaw = r.volume ? r.volume * 0.5 : null;
    // 2) 사람(비교용)
    r.qPerson = N ? N * 25 : null;

    // 3) 습기
    var G = 0; // g/일
    function src(name, g, basis) { if (g > 0) { r.sources.push({ name: name, g: g, basis: basis }); G += g; } }
    if (N) {
      var sleep = num(inp.sleepHours); if (sleep == null) sleep = 8;
      var home = num(inp.homeHours); if (home == null) home = 24;
      home = Math.max(sleep, Math.min(24, home));
      src('잠자는 동안 숨·땀', BS.asleep * sleep * N, 'BS 5250 수면 40g/h·인 × ' + sleep + '시간 × ' + N + '명');
      src('깨어 있는 동안 숨·땀', BS.seated * (home - sleep) * N, 'BS 5250 앉아서 활동 70g/h·인 × ' + (home - sleep) + '시간 × ' + N + '명');
      var sh = num(inp.showersPerDay); if (sh == null) sh = N;
      src('샤워', BS.shower * sh, 'BS 5250 샤워 600g/회 × 하루 ' + sh + '회');
      src('바닥 청소', BS.floors, 'BS 5250 200g/일');
    }
    var cd = num(inp.cookDays); if (cd == null) cd = 7;
    if (inp.cook === 'gas') { src('조리(가스)', BS.cookGas * cd / 7, 'BS 5250 가스 조리 3,000g/일 × 주 ' + cd + '일'); src('설거지', BS.dish * cd / 7, 'BS 5250 400g/일 × 주 ' + cd + '일'); }
    else if (inp.cook === 'elec') { src('조리(전기·인덕션)', BS.cookElec * cd / 7, 'BS 5250 전기 조리 2,000g/일 × 주 ' + cd + '일'); src('설거지', BS.dish * cd / 7, 'BS 5250 400g/일 × 주 ' + cd + '일'); }
    var ld = num(inp.laundryDays);
    if (ld) {
      src('세탁', BS.washClothes * ld / 7, 'BS 5250 500g/일 × 주 ' + ld + '일');
      if (inp.dry === 'indoor') src('빨래 실내 건조', BS.dryIndoor * ld / 7, 'BS 5250 실내 건조 1,500g/일 × 주 ' + ld + '일');
    }
    var pl = num(inp.plants); if (pl) src('화분', BS.plant * pl, 'BS 5250 20g/일·개 × ' + pl + '개');
    var hm = num(inp.humidMlh), hh = num(inp.humidHours);
    if (hm && hh) src('가습기', hm * hh, '입력값 ' + hm + 'ml/h × ' + hh + '시간');
    else if (inp.humidifier) r.notes.push('가습기를 쓴다고 했지만 가습량(ml/h)·사용 시간이 없어 계산에서 뺐습니다. 제품 표시 가습량을 넣으면 반영됩니다.');
    if (inp.aquarium) r.notes.push('수족관은 BS 5250 표에 값이 없어 계산에서 뺐습니다. 수면이 넓을수록 습기가 더 생깁니다.');
    r.moistureGday = G;

    // 실내 목표 습도: 사용자 목표와 유리 표면 결로 한계 중 낮은 값
    var Ti = num(inp.Ti); if (Ti == null) Ti = 22;
    var tRH = num(inp.targetRH); if (tRH == null) tRH = 50;
    r.Ti = Ti; r.targetRH = tRH; r.rhBasis = '목표 습도';
    var Ts = num(inp.Ts);
    if (Ts != null && Ts < Ti) {
      r.allowRH = allowRH(Ti, Ts);
      if (r.allowRH < tRH) { r.targetRH = Math.floor(r.allowRH); r.rhBasis = '측정한 유리 표면온도 ' + Ts + '℃의 결로 한계'; }
    }
    var oT = num(inp.outT), oRH = num(inp.outRH);
    if (oT == null || oRH == null) r.missing.push('겨울 바깥 기온·습도(지역)');
    if (G > 0 && oT != null && oRH != null) {
      r.xIn = mixRatio(Ti, r.targetRH); r.xOut = mixRatio(oT, oRH);
      var dx = r.xIn - r.xOut;
      r.qMoist = dx > 0 ? (G / 24) / (RHO * dx) : Infinity;
    } else r.qMoist = null;

    // 필요 풍량 = 셋 중 최대(사람 기준은 비교용이지만 안전측으로 포함)
    var cands = [['법정 0.5회', r.qLaw], ['사람 수', r.qPerson], ['습기 배출', r.qMoist]].filter(function (c) { return c[1] != null && isFinite(c[1]); });
    if (cands.length) { cands.sort(function (a, b) { return b[1] - a[1]; }); r.qReq = cands[0][1]; r.qReqBy = cands[0][0]; } else r.qReq = null;
    if (r.qMoist === Infinity) r.notes.push('목표 습도가 바깥 공기 습도보다 낮아 환기만으로는 맞출 수 없습니다. 목표 습도를 다시 확인하십시오.');

    // 생활 조건 메모
    if (inp.cantOpen) r.notes.push('창문 환기가 어려운 생활(맞벌이 등)이라 필요한 환기를 전부 기계환기가 맡아야 합니다. 창문 환기는 계산에 넣지 않았습니다.');
    if (inp.newborn) r.notes.push('신생아가 있으면 겨울에 창문을 열기 어렵고, 너무 건조해도 좋지 않습니다. 습도는 결로 한계 아래에서 지나치게 낮아지지 않게 맞추고, 가습기는 습도 설정(자동) 모드로 쓰십시오.');
    var W = { passive: '패시브하우스 수준 기밀: 틈으로 드나드는 공기가 거의 없어 환기를 전부 기계환기에 의존합니다. 환기장치가 멈추면 습기가 바로 쌓입니다.',
              system: '시스템창호(기밀형): 틈바람이 적어 기계환기 의존도가 높습니다. 상시 가동이 전제입니다.',
              double: '일반 이중창(미서기): 틈바람이 일부 환기를 대신하지만 양이 일정하지 않고 외풍·결로를 함께 부릅니다. 계산에는 넣지 않았습니다(안전측).',
              old: '노후 창호: 틈바람이 많아 실제 환기량은 계산보다 클 수 있으나, 그 대가로 외풍과 창 둘레 결로가 생깁니다. 창호 기밀 보강 뒤에는 기계환기 필요량이 그대로 드러납니다.' };
    if (W[inp.window]) r.notes.push(W[inp.window]);

    // 4) 설치된 제품 판정
    var P = inp.product || {};
    var st = { low: num(P.low), mid: num(P.mid), high: num(P.high) };
    r.hasProduct = !!(st.low || st.mid || st.high);
    if (r.hasProduct) {
      r.stages = [];
      var latent = (P.type === 'erv') ? num(P.latentEff) : (P.type === 'hrv' ? 0 : null);
      [['약', st.low], ['중', st.mid], ['강', st.high]].forEach(function (s) {
        if (!s[1]) return;
        var eff = s[1] * 0.9; // 법정 최저 유효환기량(표시의 90%)
        var moistEff = (latent != null) ? eff * (1 - latent / 100) : null; // 전열형은 습기 일부를 되돌림
        r.stages.push({ name: s[0], rated: s[1], eff: eff, moistEff: moistEff, ach: r.volume ? eff / r.volume : null });
      });
      if (P.type === 'erv' && latent == null) r.notes.push('전열형은 습기 일부를 실내로 되돌립니다. 제품 사양서의 습도(잠열) 교환효율을 넣으면 습기 배출 판정이 정확해집니다. 지금은 되돌림 없이 계산해 실제보다 유리하게 나옵니다.');
      if (!P.type) r.notes.push('전열형인지 현열형인지 모르면 습기 배출 판정은 되돌림 없이 계산합니다(실제보다 유리하게 나올 수 있음).');
      if (P.error) { r.verdict = 'hold'; r.verdictText = '판정 보류 — 조작기 에러'; r.verdictWhy = '에러가 떠 있으면 팬이 돌지 않을 수 있어 표시 풍량을 믿을 수 없습니다. 에러를 해결한 뒤 다시 판정하십시오.'; }
      else if (P.filter === 'clog' || P.filter === 'due') { r.verdict = 'hold'; r.verdictText = '판정 보류 — 필터 교체 필요'; r.verdictWhy = '막힌 필터는 풍량을 크게 줄입니다. 필터를 갈고 다시 판정하십시오. 아래 계산은 필터가 깨끗할 때의 값입니다.'; }
      if (r.qReq != null && r.verdict !== 'hold') {
        var ok = function (s) { var vMain = s.eff >= Math.max(r.qLaw || 0, r.qPerson || 0); var vMoist = (r.qMoist == null) ? true : ((s.moistEff != null ? s.moistEff : s.eff) >= r.qMoist); return vMain && vMoist; };
        var first = null; r.stages.forEach(function (s) { if (!first && ok(s)) first = s; });
        r.recommendStage = first ? first.name : null;
        // 풍량은 충분한데 습기 배출(전열형 되돌림 반영)이 단계를 끌어올린 경우 알림
        var mainOnly = function (s) { return s.eff >= Math.max(r.qLaw || 0, r.qPerson || 0); };
        var firstMain = null; r.stages.forEach(function (s) { if (!firstMain && mainOnly(s)) firstMain = s; });
        if (firstMain && first !== firstMain) r.notes.unshift('풍량만 보면 ' + firstMain.name + ' 단계로 충분하지만, 습기 배출' + (latent ? '(전열형이 습기 ' + latent + '%를 되돌림)' : '') + '이 모자라 ' + (first ? first.name + ' 단계가 필요합니다.' : '가장 센 단계로도 부족합니다.'));
        if (!first) { r.verdict = 'short'; r.verdictText = '부족'; r.verdictWhy = '가장 센 단계로 돌려도 계산상 필요 풍량에 못 미칩니다. 습기 발생을 줄이거나(실내 건조·가습기), 보조 환기·제습을 함께 써야 합니다.'; }
        else if (first.name === '약') { r.verdict = 'ample'; r.verdictText = '여유 있음'; r.verdictWhy = '약 단계로 계속 돌려도 필요 풍량을 채웁니다.'; }
        else if (first.name === '중') { r.verdict = 'ok'; r.verdictText = '적정'; r.verdictWhy = '중 단계로 계속 돌리면 필요 풍량을 채웁니다.'; }
        else { r.verdict = 'tight'; r.verdictText = '빠듯함'; r.verdictWhy = '강 단계로 계속 돌려야 필요 풍량을 채웁니다. 소음 때문에 끄기 쉬워 실제로는 모자랄 수 있습니다.'; }
      }
    } else if (r.qReq != null) {
      // 없는 집: 적정(중) 단계 표시 풍량 권장치 — 유효 90%를 감안해 50CMH 단위로 올림
      r.recommendRated = Math.ceil((r.qReq / 0.9) / 50) * 50;
      r.verdict = 'none'; r.verdictText = '권장 제품 급';
      r.verdictWhy = '중(적정) 단계 풍량이 ' + r.recommendRated + 'CMH 이상인 제품이 맞습니다(표시 풍량의 90%를 실제 풍량으로 보고 계산).';
    }
    r.missingAll = r.missing.length > 0;
    return r;
  }

  // ── 지난겨울(12~2월) 일평균 기온·상대습도 → 평균 (Open-Meteo archive)
  async function fetchWinterClimate(lat, lon) {
    var now = new Date(), y = now.getFullYear(), m = now.getMonth() + 1;
    var endY = (m >= 3) ? y : y - 1;
    var start = (endY - 1) + '-12-01', end = endY + '-02-28';
    var url = 'https://archive-api.open-meteo.com/v1/archive?latitude=' + lat + '&longitude=' + lon +
      '&daily=temperature_2m_mean,relative_humidity_2m_mean&timezone=Asia%2FSeoul&start_date=' + start + '&end_date=' + end;
    var res = await fetch(url); if (!res.ok) throw new Error('기상자료 조회 실패');
    var j = await res.json(), D = j.daily; if (!D || !D.time) throw new Error('기상자료 없음');
    var sT = 0, sX = 0, sRH = 0, n = 0, minT = Infinity;
    for (var i = 0; i < D.time.length; i++) {
      var T = D.temperature_2m_mean[i], RH = D.relative_humidity_2m_mean[i];
      if (T == null || RH == null) continue; sT += T; sRH += RH; sX += mixRatio(T, RH); n++; if (T < minT) minT = T;
    }
    if (!n) throw new Error('기상자료 없음');
    var meanT = sT / n, meanX = sX / n;
    // 평균 절대습도에 해당하는 평균기온에서의 상대습도(계산 일관성용)
    var e = meanX * 101.325 / (622 + meanX), rh = Math.min(100, e / es(meanT) * 100);
    return { period: start + ' ~ ' + end, meanT: meanT, meanRH: rh, meanRHraw: sRH / n, meanX: meanX, minDailyT: minT, days: n };
  }

  async function geocode(addr) {
    var words = String(addr || '').replace(/[○◯]+/g, '').split(/\s+/).filter(Boolean);
    while (words.length) {
      var q = words.join(' ');
      try {
        var r = await fetch('https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=kr&accept-language=ko&q=' + encodeURIComponent(q));
        if (r.ok) { var j = await r.json(); if (j && j[0]) return { lat: +j[0].lat, lon: +j[0].lon, matched: q }; }
      } catch (e) {}
      words.pop();
    }
    return null;
  }

  // 결과 → 사람이 읽는 요약(소견서·AI 프롬프트용)
  function summaryText(r) {
    var f0 = function (x) { return (x == null || !isFinite(x)) ? '-' : Math.round(x); };
    var L = [];
    L.push('- 필요 풍량: 법정 0.5회 ' + f0(r.qLaw) + 'CMH / 사람 수(비교) ' + f0(r.qPerson) + 'CMH / 습기 배출 ' + f0(r.qMoist) + 'CMH → 기준 ' + f0(r.qReq) + 'CMH(' + (r.qReqBy || '-') + ')');
    if (r.moistureGday) L.push('- 하루 수분 발생량(BS 5250 Table D.6 기준 계산): 약 ' + (r.moistureGday / 1000).toFixed(1) + 'kg — ' + r.sources.map(function (s) { return s.name + ' ' + (s.g / 1000).toFixed(2) + 'kg'; }).join(', '));
    L.push('- 실내 ' + r.Ti + '℃·습도 상한 ' + r.targetRH + '%(' + r.rhBasis + ')' + (r.xOut != null ? ', 겨울 바깥 절대습도 ' + r.xOut.toFixed(1) + 'g/kg' : ''));
    if (r.hasProduct) L.push('- 설치 제품: ' + r.stages.map(function (s) { return s.name + ' ' + s.rated + 'CMH(유효 ' + f0(s.eff) + ')'; }).join(' / ') + ' → 판정: ' + (r.verdictText || '-') + (r.recommendStage ? ' (권장 운전 ' + r.recommendStage + ' 단계)' : ''));
    else if (r.recommendRated) L.push('- 환기장치 없음 → 권장: 중 단계 ' + r.recommendRated + 'CMH 이상 제품');
    r.notes.forEach(function (n) { L.push('- 참고: ' + n); });
    return L.join('\n');
  }

  root.ERV = { compute: compute, fetchWinterClimate: fetchWinterClimate, geocode: geocode, summaryText: summaryText, mixRatio: mixRatio, allowRH: allowRH, BS: BS, VERSION: ENGINE_VERSION };
})(typeof window !== 'undefined' ? window : globalThis);
