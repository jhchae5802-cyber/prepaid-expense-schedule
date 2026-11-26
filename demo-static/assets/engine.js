// 선급비용 명세서 계산 엔진 (브라우저용)
// webapp/schedule_core.py 의 Python 검증 로직을 그대로 옮긴 것.
// 규칙: 분 단위 기간, 기준일 24:00, 잔여비율 소수 5자리 반올림, 잔액 원 단위 반올림,
//       월별 비용 = 전월말 잔액 - 당월말 잔액, 회계일자 이전 월 잔액 0, 원장 대사는 각 행의 결산일자 기준.
(function (global) {
  'use strict';
  var DAY = 86400000, MIN = 60000;

  // 'yyyymmddhhmm' / 'yyyymmdd' / 'yyyy-mm-dd' → UTC ms
  function pdt(v) {
    var s = String(v).replace(/[-.: ]/g, '');
    return Date.UTC(+s.slice(0, 4), +s.slice(4, 6) - 1, +s.slice(6, 8), +(s.slice(8, 10) || 0), +(s.slice(10, 12) || 0));
  }
  function iso(ms) { return new Date(ms).toISOString().slice(0, 10); }
  function isoMin(ms) { return new Date(ms).toISOString().slice(0, 16).replace('T', ' '); }
  function monthEnd(y, m) { return Date.UTC(y, m, 0); }          // m: 1~12
  // Excel ROUND (0에서 먼 쪽으로 반올림) — Python r0 와 같은 연산 순서
  function r0(x, k) { var f = Math.pow(10, k || 0); return (Math.floor(Math.abs(x) * f + 0.5) / f) * (x >= 0 ? 1 : -1); }
  function pyRound(x) { var f = Math.floor(x), d = x - f; return d > 0.5 ? f + 1 : d < 0.5 ? f : (f % 2 === 0 ? f : f + 1); }

  function bal(x, d) {                                          // d(날짜 ms) 24:00 시점 잔액
    var J = (x.I - x.H) / MIN;
    var rem = Math.max(0, (x.I - Math.max(d + DAY, x.H)) / MIN);
    return r0(x.amt * r0(pyRound(rem) / pyRound(J), 5));
  }

  function build(src, baseISO) {
    var rows = src.rows.map(function (r, k) {
      return { r: k + 2, amt: +r.amt, H: pdt(r.start), I: pdt(r.end), pay: r.date ? pdt(String(r.date).slice(0, 8)) : null,
               acct: r.acct, close: r.close ? pdt(String(r.close).slice(0, 8)) : null, ledger: r.ledger,
               voucher: r.voucher || '', vendor: r.vendor || '', memo: r.memo || '' };
    });
    var closes = rows.filter(function (x) { return x.close != null; }).map(function (x) { return x.close; });
    var base, baseAuto = false;
    if (baseISO) base = pdt(baseISO);
    else if (closes.length) {                                    // 결산일자 최빈값(동률이면 늦은 날)
      var cnt = {}; closes.forEach(function (c) { cnt[c] = (cnt[c] || 0) + 1; });
      base = Object.keys(cnt).map(Number).sort(function (a, b) { return cnt[b] - cnt[a] || b - a; })[0]; baseAuto = true;
    }
    var accts = []; rows.forEach(function (x) { if (accts.indexOf(x.acct) < 0) accts.push(x.acct); });
    var minH = new Date(Math.min.apply(null, rows.map(function (x) { return x.H; })));
    var maxI = new Date(Math.max.apply(null, rows.map(function (x) { return x.I; })));
    var months = [], y = minH.getUTCFullYear(), m = minH.getUTCMonth() + 1;
    while (y < maxI.getUTCFullYear() || (y === maxI.getUTCFullYear() && m <= maxI.getUTCMonth() + 1)) {
      months.push(monthEnd(y, m)); m++; if (m === 13) { y++; m = 1; }
    }
    var NM = months.length, zeros = function () { return months.map(function () { return 0; }); };
    var hasLed = rows.some(function (x) { return x.ledger != null; });
    var exp = zeros(), mb = zeros(), tb = 0, accB = {}, accE = {}, accM = {}, accA = {}, bad = [], ledBad = [], items = [];
    accts.forEach(function (a) { accB[a] = 0; accE[a] = zeros(); accM[a] = zeros(); accA[a] = 0; });

    rows.forEach(function (x, k) {
      var mv = months.map(function (d) { var prev = Date.UTC(new Date(d).getUTCFullYear(), new Date(d).getUTCMonth(), 0); return bal(x, prev) - bal(x, d); });
      var mbal = months.map(function (d) { return (x.pay != null && d < x.pay) ? 0 : bal(x, d); });
      var s = mv.reduce(function (a, b) { return a + b; }, 0);
      if (Math.abs(s - x.amt) > 0.5) bad.push(x.r);
      for (var i = 0; i < NM; i++) { exp[i] += mv[i]; mb[i] += mbal[i]; accE[x.acct][i] += mv[i]; accM[x.acct][i] += mbal[i]; }
      accA[x.acct] += x.amt;
      var J = pyRound((x.I - x.H) / MIN), rem = pyRound(Math.max(0, (x.I - Math.max(base + DAY, x.H)) / MIN));
      var b = bal(x, base); tb += b; accB[x.acct] += b;
      var led = null, recal = null, diff = null;
      if (hasLed && x.ledger != null) {
        led = Math.abs(+x.ledger); recal = bal(x, x.close != null ? x.close : base); diff = recal - led;
        if (Math.abs(diff) > 0.5) ledBad.push({ row: x.r, no: k + 1, diff: diff });
      }
      items.push({ no: k + 1, row: x.r, voucher: x.voucher, pay: x.pay != null ? iso(x.pay) : '', acct: x.acct, vendor: x.vendor, memo: x.memo,
                   amt: x.amt, start: isoMin(x.H), end: isoMin(x.I), total_min: J, rem_min: rem, ratio: J ? r0(rem / J, 5) : 0,
                   expense: mv, balance: mbal, cum: x.amt - b, base_bal: b, close: x.close != null ? iso(x.close) : '',
                   ledger: led, recalc: recal, diff: diff });
    });

    var warnings = [];
    if (closes.length && closes.indexOf(base) < 0) {
      warnings.push('원자료에 기준일(' + iso(base) + ') 결산일자가 없습니다. 결산일자는 ' + iso(Math.max.apply(null, closes)) + '까지만 있습니다. 해당 월 원자료가 맞는지 확인하세요.');
    }
    var dist = {}; closes.forEach(function (c) { var k = iso(c); dist[k] = (dist[k] || 0) + 1; });
    var distSorted = {}; Object.keys(dist).sort().forEach(function (k) { distSorted[k] = dist[k]; });
    var bd = new Date(base);
    return {
      sheet: src.sheet, first: 2, last: rows.length + 1, n: rows.length, base: iso(base), base_auto: baseAuto,
      filename: "선급비용명세서('" + String(bd.getUTCFullYear()).slice(2) + '.' + (bd.getUTCMonth() + 1) + ").xlsx", has_ledger: hasLed,
      months: months.map(function (d) { return iso(d).slice(0, 7); }),
      total_balance: tb, total_amount: rows.reduce(function (a, x) { return a + x.amt; }, 0),
      by_account: accts.map(function (a) { return { acct: a, amount: accA[a], balance: accB[a], expense: accE[a], month_balance: accM[a] }; }),
      month_expense: exp, month_balance: mb, check_sum_ok: !bad.length, check_sum_bad: bad,
      ledger_ok: hasLed ? !ledBad.length : null, ledger_bad: ledBad, close_dist: distSorted, warnings: warnings, items: items
    };
  }

  global.PrepaidEngine = { build: build };
})(window);
