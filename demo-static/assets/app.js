// 화면 동작. 원본 웹앱(webapp/templates/index.html)의 화면을 서버 없이 돌도록 바꾼 것.
// - /api/build      → PrepaidEngine.build(SAMPLE, 기준일) 브라우저 계산
// - 파일 업로드      → 전시용 시뮬레이션 (파일을 읽지 않고 샘플 원자료 사용)
// - /api/download   → 미리 만들어 둔 샘플 명세서 xlsx (2026년 1~12월 기준)
// - /api/save       → 전시용 시뮬레이션 (아무것도 저장하지 않음)
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var fmt = function (v) { return v == null || v === '' ? '' : (Math.round(v) === 0 ? '-' : Math.round(v).toLocaleString('ko-KR')); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var sum = function (a) { return a.reduce(function (s, v) { return s + v; }, 0); };
  var R = null, tab = 'month';

  $('#srcName').textContent = SAMPLE.source_name;

  // ── 업로드: 전시용 시뮬레이션 ──
  var drop = $('#drop'), fi = $('#file');
  drop.addEventListener('click', function (e) { if (e.target.tagName !== 'A') fi.click(); });
  drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  fi.addEventListener('change', function () { simUpload(fi.files[0]); fi.value = ''; });
  ['dragover', 'dragenter'].forEach(function (e) { drop.addEventListener(e, function (ev) { ev.preventDefault(); drop.classList.add('hover'); }); });
  ['dragleave', 'drop'].forEach(function (e) { drop.addEventListener(e, function (ev) { ev.preventDefault(); drop.classList.remove('hover'); }); });
  drop.addEventListener('drop', function (ev) { simUpload(ev.dataTransfer.files[0]); });
  function simUpload(f) {
    if (!f) return;
    clearMsgs();
    msg('<span class="sim">전시용 시뮬레이션</span> 선택한 파일 “' + esc(f.name) + '”은 읽거나 전송하지 않았습니다. 데모에서는 샘플 원자료로 계산합니다.', 'simm');
    var m = f.name.match(/'(\d{2})\.(\d{1,2})\)/);
    if (m) $('#month').value = '20' + m[1] + '-' + ('0' + m[2]).slice(-2);
    run(true);
  }

  function monthEnd(v) { var p = v.split('-').map(Number); var d = new Date(Date.UTC(p[0], p[1], 0)).getUTCDate(); return v + '-' + ('0' + d).slice(-2); }
  function msg(html, cls) { $('#msgs').insertAdjacentHTML('beforeend', '<div class="msg ' + cls + '">' + html + '</div>'); }
  function clearMsgs(keepSim) {
    Array.prototype.forEach.call($('#msgs').children, function (e) { if (!(keepSim && e.classList.contains('simm'))) e.remove(); });
  }

  function run(keepSim) {
    clearMsgs(keepSim);
    try {
      R = PrepaidEngine.build(SAMPLE, $('#month').value ? monthEnd($('#month').value) : null);
      R.source_name = SAMPLE.source_name;
      if (!$('#month').value) $('#month').value = R.base.slice(0, 7);
      render();
    } catch (e) { msg('계산 중 오류: ' + esc(e.message), 'err'); }
  }
  $('#run').addEventListener('click', function () { run(false); });
  $('#month').addEventListener('change', function () { run(false); });

  function render() {
    $('#result').classList.remove('hidden');
    R.warnings.forEach(function (w) { msg('⚠ ' + esc(w), 'warn'); });
    var ok1 = R.check_sum_ok, ok2 = R.ledger_ok;
    var tiles = R.by_account.map(function (a) {
      return '<div class="kpi"><div class="t">' + esc(a.acct) + '</div><div class="v">' + fmt(a.balance) + '</div><div class="s">발생금액 ' + fmt(a.amount) + '</div></div>';
    }).join('');
    $('#kpis').innerHTML =
      '<div class="kpi main"><div class="t">기준일 선급비용 잔액 · ' + R.base + (R.base_auto ? ' (자동)' : '') + '</div><div class="v">' + fmt(R.total_balance) +
      '</div><div class="s">' + esc(R.source_name) + ' · ' + R.n + '건 · 발생금액 ' + fmt(R.total_amount) + '</div></div>' + tiles +
      '<div class="kpi"><div class="t">검증</div><div style="margin-top:4px">월별비용 합계 = 발생금액 <span class="badge ' + (ok1 ? 'ok' : 'bad') + '">' +
      (ok1 ? '전건 일치' : '불일치 ' + R.check_sum_bad.length + '건') + '</span></div>' +
      (R.has_ledger ? '<div>원장 대사(결산일자 기준) <span class="badge ' + (ok2 ? 'ok' : 'bad') + '">' + (ok2 ? '전건 일치' : '차이 ' + R.ledger_bad.length + '건') + '</span></div>'
                    : '<div class="s">발생계산금액 열 없음 — 원장 대사 생략</div>') + '</div>';
    var dl = $('#dl'), has = SAMPLE.sample_bases.indexOf(R.base) >= 0;
    dl.classList.toggle('disabled', !has);
    if (has) {
      dl.href = './assets/samples/schedule-' + R.base.slice(0, 7) + '.xlsx';
      dl.setAttribute('download', R.filename);
      dl.textContent = '엑셀 다운로드 · ' + R.filename;
      dl.removeAttribute('aria-disabled');
    } else {
      dl.href = '#'; dl.removeAttribute('download'); dl.setAttribute('aria-disabled', 'true');
      dl.textContent = '샘플 엑셀은 2026년 1~12월 기준만 제공';
    }
    draw();
  }

  Array.prototype.forEach.call(document.querySelectorAll('#tabs button'), function (b) {
    b.addEventListener('click', function () {
      Array.prototype.forEach.call(document.querySelectorAll('#tabs button'), function (x) { x.classList.toggle('on', x === b); });
      tab = b.dataset.t; draw();
    });
  });

  function mcls(i) { var m = R.months[i], bm = R.base.slice(0, 7); return m === bm ? 'cur' : (m < bm ? 'past' : ''); }
  function mlabel(m) { return m.slice(0, 4) + '.' + m.slice(5); }

  function draw() {
    var T = $('#tbl'), h = '';
    var legendMonths = '<i style="background:var(--past)"></i>기준일 이전 월<i style="background:var(--cur)"></i>기준월';
    var A = R.by_account;
    if (tab === 'month') {
      h = '<tr><th>월</th><th>발생 비용(합계)</th>' + A.map(function (a) { return '<th>' + esc(a.acct) + '<br>비용</th>'; }).join('') +
          '<th>월말 잔액(합계)</th>' + A.map(function (a) { return '<th>' + esc(a.acct) + '<br>잔액</th>'; }).join('') + '</tr>';
      R.months.forEach(function (m, i) {
        var c = mcls(i);
        h += '<tr><td class="c ' + c + '">' + mlabel(m) + '</td><td class="n ' + c + '">' + fmt(R.month_expense[i]) + '</td>' +
             A.map(function (a) { return '<td class="n ' + c + '">' + fmt(a.expense[i]) + '</td>'; }).join('') +
             '<td class="n ' + c + '"><b>' + fmt(R.month_balance[i]) + '</b></td>' +
             A.map(function (a) { return '<td class="n ' + c + '">' + fmt(a.month_balance[i]) + '</td>'; }).join('') + '</tr>';
      });
      h += '<tr class="tot"><td class="c">합계</td><td class="n">' + fmt(sum(R.month_expense)) + '</td>' +
           A.map(function (a) { return '<td class="n">' + fmt(sum(a.expense)) + '</td>'; }).join('') + '<td></td>' + A.map(function () { return '<td></td>'; }).join('') + '</tr>';
      $('#legend').innerHTML = '월말 잔액은 회계일자(지급) 이전 월을 0으로 둡니다.' + legendMonths;
    } else if (tab === 'items') {
      h = '<tr><th>No</th><th>전표번호</th><th>회계일자</th><th>계정과목명</th><th>거래처명</th><th>적요</th><th>발생금액</th><th>시작시간</th><th>종료시간</th><th>총기간(분)</th><th>잔여기간(분)</th><th>잔여비율</th><th>기준일까지 비용 누계</th><th>기준일 잔액</th></tr>';
      R.items.forEach(function (x) {
        h += '<tr><td class="c">' + x.no + '</td><td>' + esc(x.voucher) + '</td><td class="c">' + x.pay + '</td><td>' + esc(x.acct) + '</td><td>' + esc(x.vendor) +
             '</td><td class="memo" title="' + esc(x.memo) + '">' + esc(x.memo) + '</td><td class="n">' + fmt(x.amt) + '</td><td class="c">' + x.start + '</td><td class="c">' + x.end +
             '</td><td class="n">' + x.total_min.toLocaleString() + '</td><td class="n">' + x.rem_min.toLocaleString() + '</td><td class="c">' + x.ratio.toFixed(5) +
             '</td><td class="n">' + fmt(x.cum) + '</td><td class="n"><b>' + fmt(x.base_bal) + '</b></td></tr>';
      });
      h += '<tr class="tot"><td colspan="6" class="c">합계</td><td class="n">' + fmt(R.total_amount) + '</td><td colspan="5"></td><td class="n">' +
           fmt(R.total_amount - R.total_balance) + '</td><td class="n">' + fmt(R.total_balance) + '</td></tr>';
      $('#legend').innerHTML = "원자료 시트 '" + esc(R.sheet) + "' " + R.first + '~' + R.last + '행. 잔여비율 = ROUND(잔여분 ÷ 총분, 5), 잔액 = ROUND(발생금액 × 잔여비율, 0).';
    } else if (tab === 'matrix' || tab === 'balmatrix') {
      var key = tab === 'matrix' ? 'expense' : 'balance', isE = key === 'expense';
      h = '<tr><th class="stick">No</th><th class="stick2">계정과목명</th><th>거래처명</th><th>발생금액</th>' +
          R.months.map(function (m, i) { return '<th class="' + (mcls(i) === 'cur' ? 'cur' : '') + '">' + mlabel(m) + '</th>'; }).join('') + (isE ? '<th>합계</th>' : '') + '</tr>';
      R.items.forEach(function (x) {
        h += '<tr><td class="c stick">' + x.no + '</td><td class="stick2">' + esc(x.acct) + '</td><td>' + esc(x.vendor) + '</td><td class="n">' + fmt(x.amt) + '</td>' +
             x[key].map(function (v, i) { return '<td class="n ' + mcls(i) + '">' + fmt(v) + '</td>'; }).join('') + (isE ? '<td class="n"><b>' + fmt(sum(x.expense)) + '</b></td>' : '') + '</tr>';
      });
      var tots = isE ? R.month_expense : R.month_balance;
      h += '<tr class="tot"><td class="stick"></td><td class="stick2">합계</td><td></td><td class="n">' + fmt(R.total_amount) + '</td>' +
           tots.map(function (v) { return '<td class="n">' + fmt(v) + '</td>'; }).join('') + (isE ? '<td class="n">' + fmt(sum(tots)) + '</td>' : '') + '</tr>';
      $('#legend').innerHTML = (isE ? '월별 비용 = 전월말 잔액 − 당월말 잔액. 건별 합계는 발생금액과 같습니다.' : '각 월말 24:00 기준 잔액. 회계일자 이전 월은 0.') + legendMonths;
    } else {
      h = '<tr><th>No</th><th>원자료 행</th><th>계정과목명</th><th>거래처명</th><th>적요</th><th>원장 결산일자</th><th>원장 잔액<br>(발생계산금액)</th><th>결산일자 기준<br>재계산 잔액</th><th>차이</th></tr>';
      R.items.forEach(function (x) {
        h += '<tr><td class="c">' + x.no + '</td><td class="c">' + x.row + '</td><td>' + esc(x.acct) + '</td><td>' + esc(x.vendor) + '</td><td class="memo" title="' + esc(x.memo) + '">' +
             esc(x.memo) + '</td><td class="c">' + (x.close || '(기준일)') + '</td><td class="n">' + fmt(x.ledger) + '</td><td class="n">' + fmt(x.recalc) +
             '</td><td class="n ' + (x.diff && Math.abs(x.diff) > 0.5 ? 'bad' : '') + '">' + fmt(x.diff) + '</td></tr>';
      });
      var s = function (k) { return R.items.reduce(function (a, x) { return a + (x[k] || 0); }, 0); };
      h += '<tr class="tot"><td colspan="6" class="c">합계</td><td class="n">' + fmt(s('ledger')) + '</td><td class="n">' + fmt(s('recalc')) + '</td><td class="n">' + fmt(s('diff')) + '</td></tr>';
      var dist = Object.keys(R.close_dist).map(function (d) { return d + ' ' + R.close_dist[d] + '건'; }).join(', ');
      $('#legend').innerHTML = '원자료에는 이전 결산에서 넘어온 행이 섞여 있으므로 각 행의 결산일자 시점 잔액과 비교합니다. 결산일자 분포: ' + esc(dist);
    }
    T.innerHTML = h;
  }

  // ── 작업 폴더 저장: 전시용 시뮬레이션 ──
  $('#save').addEventListener('click', function () {
    Array.prototype.forEach.call(document.querySelectorAll('#msgs .ok'), function (e) { e.remove(); });
    msg('<span class="sim">전시용 시뮬레이션</span> 실제 도구에서는 원자료 폴더에 ' + esc(R.filename) +
        ' 파일을 저장합니다(같은 이름이 있으면 덮어쓸지 먼저 묻습니다). 데모에서는 아무것도 저장하지 않았습니다.', 'ok');
  });

  run(false);
})();
