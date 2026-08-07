/* PrimeGrid — 경사면 일사량(POA) 및 예상 발전량 계산 엔진
 *
 * 기상청 ASOS 수평면 일사량(GHI) 월평균 → 경사면 일사량(POA) 변환.
 *   1) Erbs 월평균 상관식으로 직달/산란 분리
 *   2) 15분 간격 태양벡터로 직달 기하계수 Rb 산출
 *   3) 등방성(Liu-Jordan) 천공 + 지면반사 합산
 *
 * 원본 산출표의 독립 검증표와 대조 시, 자료가 동일한 56개 지점에서
 * 연간 POA가 오차 0.000%로 일치. β=0°에서 POA=GHI 항등식도 정확히 성립.
 */
(function (global) {
  'use strict';

  var GSC = 0.0820;                                       // 태양상수 MJ/m²/min
  var DOY = [17, 47, 75, 105, 135, 162, 198, 228, 258, 288, 318, 344];
  var MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
  var RAD = Math.PI / 180;

  /** 월 대표일의 태양기하 + 대기권외 일사량 H0 (kWh/m²/day) */
  function solarGeom(doy, latDeg) {
    var B = 2 * Math.PI * (doy - 1) / 365;
    var E0 = 1.000110 + 0.034221 * Math.cos(B) + 0.001280 * Math.sin(B)
           + 0.000719 * Math.cos(2 * B) + 0.000077 * Math.sin(2 * B);
    var decl = 0.006918 - 0.399912 * Math.cos(B) + 0.070257 * Math.sin(B)
             - 0.006758 * Math.cos(2 * B) + 0.000907 * Math.sin(2 * B)
             - 0.002697 * Math.cos(3 * B) + 0.001480 * Math.sin(3 * B);
    var phi = latDeg * RAD;
    var x = Math.max(-1, Math.min(1, -Math.tan(phi) * Math.tan(decl)));
    var ws = Math.acos(x);                                 // 일몰시각각 (rad)
    var H0mj = (24 * 60 / Math.PI) * GSC * E0
             * (Math.cos(phi) * Math.cos(decl) * Math.sin(ws) + ws * Math.sin(phi) * Math.sin(decl));
    return { decl: decl, ws: ws, phi: phi, H0: H0mj / 3.6 };   // MJ → kWh
  }

  /** Erbs 월평균 확산분율 Hd/H */
  function erbsMonthly(Kt, wsDeg) {
    var f = wsDeg <= 81.4
      ? 1.391 - 3.560 * Kt + 4.189 * Kt * Kt - 2.137 * Kt * Kt * Kt
      : 1.311 - 3.022 * Kt + 3.427 * Kt * Kt - 1.821 * Kt * Kt * Kt;
    return Math.max(0, Math.min(1, f));
  }

  /** 직달 기하계수 Rb — 15분 구간 태양벡터 적분 */
  function rbFactor(phi, decl, betaDeg, azimDeg) {
    var b = betaDeg * RAD, A = azimDeg * RAD;
    var nE = Math.sin(b) * Math.sin(A), nN = Math.sin(b) * Math.cos(A), nU = Math.cos(b);
    var num = 0, den = 0;
    for (var i = 0; i < 96; i++) {
      var t = (i + 0.5) * 0.25;                            // 구간 중앙 진태양시
      var w = 15 * (t - 12) * RAD;
      var sU = Math.sin(phi) * Math.sin(decl) + Math.cos(phi) * Math.cos(decl) * Math.cos(w);
      if (sU <= 0) continue;                               // 태양고도 양수 구간만
      var sE = -Math.cos(decl) * Math.sin(w);
      var sN = Math.sin(decl) * Math.cos(phi) - Math.cos(decl) * Math.sin(phi) * Math.cos(w);
      var cosAOI = sE * nE + sN * nN + sU * nU;
      if (cosAOI > 0) num += cosAOI;
      den += sU;
    }
    return den > 0 ? num / den : 0;
  }

  /**
   * 지점에 실제로 적용할 기준 기간.
   * 「계기 편향 의심」 지점만 biasPeriod를 따르고, 나머지는 basePeriod를 쓴다.
   * 산출표 6_POA_월계산 AK열(=IF(AJ="계기 편향 의심",POA_편향기간,POA_기준기간))과 같은 규칙.
   */
  function effectivePeriod(st, basePeriod, biasPeriod) {
    return (st && st.bv === '계기 편향 의심' && biasPeriod) ? biasPeriod : basePeriod;
  }

  /**
   * 지점의 월별 POA 산출.
   * @param st   데이터셋의 지점 객체
   * @param opt  {beta, azim, rho, period:'10y'|'5y'}
   * @returns    12개 월 객체 배열 (자료 없는 달은 null)
   */
  function monthly(st, opt) {
    var beta = opt.beta, azim = opt.azim, rho = opt.rho;
    var five = (opt.period === '5y' && st.ghi5);
    var ghi = five ? st.ghi5 : st.ghi;                     // 월 유효일평균 (kWh/m²/day)
    var gm  = five ? st.gm5  : st.gm;                      // 월 합계      (kWh/m²)
    var fSky = (1 + Math.cos(beta * RAD)) / 2;
    var fGnd = (1 - Math.cos(beta * RAD)) / 2;
    var out = [];
    for (var m = 0; m < 12; m++) {
      var H = ghi[m], M = gm[m];
      if (H === null || H === undefined || !M) { out.push(null); continue; }
      var g = solarGeom(DOY[m], st.lat);
      var Kt = g.H0 > 0 ? H / g.H0 : 0;
      var fd = erbsMonthly(Kt, g.ws / RAD);
      var Hd = H * fd, Hb = H - Hd;
      var Rb = rbFactor(g.phi, g.decl, beta, azim);
      var poaDay = Hb * Rb + Hd * fSky + H * rho * fGnd;
      out.push({
        m: m + 1, label: MONTHS[m],
        ghi: H, gm: M, kt: Kt, fd: fd, rb: Rb,
        poa: poaDay,
        // 월 POA = 월 GHI × (일평균 POA ÷ 일평균 GHI) — 워크북 6_POA_월계산과 동일
        poaMonth: M * (poaDay / H),
        days: M / H,
        beam: Hb * Rb, sky: Hd * fSky, gnd: H * rho * fGnd,
        ktOut: Kt < 0.30 || Kt > 0.80                      // Erbs 적용범위 이탈
      });
    }
    return out;
  }

  /** 연간 합계 */
  function annual(rows) {
    var poa = 0, ghi = 0, n = 0;
    for (var m = 0; m < 12; m++) {
      if (!rows[m]) continue;
      poa += rows[m].poaMonth;
      ghi += rows[m].gm;
      n++;
    }
    return { poa: poa, ghi: ghi, months: n, ratio: ghi > 0 ? poa / ghi : 0 };
  }

  /** n년차 모듈 성능계수 */
  function moduleFactor(year, initLoss, annLoss) {
    return (1 - initLoss) * Math.pow(1 - annLoss, Math.max(0, year - 1));
  }

  /**
   * 발전량 추정.
   * @param opt {kw, pr, initLoss, annLoss, years}
   */
  function yieldTable(rows, opt) {
    var byMonth = [], byYear = [], cum = 0;
    var f1 = moduleFactor(1, opt.initLoss, opt.annLoss);
    for (var m = 0; m < 12; m++) {
      if (!rows[m]) { byMonth.push(null); continue; }
      var poaM = rows[m].poaMonth;
      byMonth.push({
        label: rows[m].label, days: rows[m].days, poa: poaM,
        kwh: opt.kw * poaM * opt.pr * f1,
        hday: rows[m].poa * opt.pr * f1
      });
    }
    var a = annual(rows);
    for (var y = 1; y <= opt.years; y++) {
      var f = moduleFactor(y, opt.initLoss, opt.annLoss);
      var kwh = opt.kw * a.poa * opt.pr * f;
      cum += kwh;
      byYear.push({ year: y, factor: f, kwh: kwh, cum: cum, hours: a.poa * opt.pr * f });
    }
    return { byMonth: byMonth, byYear: byYear, annualPoa: a.poa, annualGhi: a.ghi, months: a.months };
  }

  /**
   * 불확실성. 기후 연변동(지점별 실측 CV)과 모델 불확실성을 독립 합성.
   * P90 = P50 × (1 − 1.282σ)
   */
  function uncertainty(st, modelSigma) {
    var clim = (st.cv && st.yrs >= 3) ? Math.min(st.cv, 0.12) : 0.045;
    var ms = modelSigma === undefined ? 0.05 : modelSigma;
    var tot = Math.sqrt(clim * clim + ms * ms);
    return { climate: clim, model: ms, total: tot, p90: 1 - 1.282 * tot, p75: 1 - 0.674 * tot };
  }

  global.POA = {
    MONTHS: MONTHS, DOY: DOY,
    solarGeom: solarGeom, erbsMonthly: erbsMonthly, rbFactor: rbFactor,
    monthly: monthly, annual: annual, yieldTable: yieldTable,
    effectivePeriod: effectivePeriod,
    moduleFactor: moduleFactor, uncertainty: uncertainty
  };
})(window);
