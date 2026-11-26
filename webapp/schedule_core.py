"""선급비용 명세서 생성 엔진 (웹앱용)

스킬 prepaid-expense-schedule 의 scripts/build_schedule.py 를 함수로 옮긴 것이다.
Excel 수식과 계산 규칙(분 단위, 기준일 24:00, 비율 5자리 반올림, 원 단위 반올림)은 원본과 같다.
build() 는 명세서 xlsx 를 저장하고, Python 독립 재계산 결과를 dict 로 돌려준다.
"""
import calendar, datetime as dt
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
from openpyxl.utils import quote_sheetname

COLS = {
    'close': ['결산일자'], 'voucher': ['전표번호'], 'date': ['회계일자', '전표일자'],
    'acct': ['계정과목명', '계정명'], 'vendor': ['거래처명'], 'memo': ['적요'],
    'krw': ['원화금액'], 'amt': ['발생금액'], 'start': ['시작시간', '시작일시', '시작일'],
    'end': ['종료시간', '종료일시', '종료일'], 'ledger': ['발생계산금액'], 'longterm': ['장기여부'],
}
REQUIRED = ['acct', 'start', 'end']


class ScheduleError(Exception):
    pass


def find_source(wb, name):
    cands = [wb[name]] if name else wb.worksheets
    for ws in cands:
        hdr = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
        m = {k: next((hdr[h] for h in v if h in hdr), None) for k, v in COLS.items()}
        if all(m[k] for k in REQUIRED) and (m['krw'] or m['amt']):
            return ws, m
    raise ScheduleError('원자료 시트를 찾지 못했습니다. 1행에 계정과목명/시작시간/종료시간/발생금액(또는 원화금액) 헤더가 필요합니다.')


def pdt(v):
    if isinstance(v, dt.datetime): return v
    if isinstance(v, dt.date): return dt.datetime.combine(v, dt.time())
    s = str(v).strip().replace('-', '').replace('.', '').replace(':', '').replace(' ', '')
    return dt.datetime(int(s[:4]), int(s[4:6]), int(s[6:8]), int(s[8:10] or 0), int(s[10:12] or 0))


def r0(x, k=0):  # Excel ROUND (half away from zero)
    f = 10 ** k
    return (int(abs(x) * f + 0.5) / f) * (1 if x >= 0 else -1)


def out_filename(base):
    return f"선급비용명세서('{base:%y}.{base.month}).xlsx"


def build(src_path, out_path=None, base_date=None, sheet=None, out_sheet='명세서'):
    """명세서를 만들어 out_path 에 저장(None 이면 저장 생략)하고 검증 결과를 반환한다."""
    wb = openpyxl.load_workbook(src_path)
    sh, cm = find_source(wb, sheet)
    S = quote_sheetname(sh.title)
    amt_col = cm['krw'] or cm['amt']

    first = 2; last = 1
    for r in range(2, sh.max_row + 1):
        if sh.cell(r, cm['start']).value not in (None, '') and sh.cell(r, amt_col).value not in (None, ''):
            last = r
        elif last >= first:
            break
    n = last - first + 1
    if n <= 0: raise ScheduleError('데이터 행이 없습니다.')

    cl_all = []
    if cm['close']:
        cl_all = [pdt(sh.cell(r, cm['close']).value).date() for r in range(first, last + 1) if sh.cell(r, cm['close']).value]
    if base_date:
        base = base_date if isinstance(base_date, dt.date) else dt.date.fromisoformat(base_date)
        base_auto = False
    elif cl_all:
        base = max(set(cl_all), key=lambda d: (cl_all.count(d), d)); base_auto = True
    else:
        t = dt.date.today(); base = t.replace(day=1) - dt.timedelta(1); base_auto = True

    if cm['longterm'] and not sh.data_validations.dataValidation:
        for o in wb.worksheets:
            if o is sh: continue
            vals = [c.value for c in o[L(cm['longterm'])] if c.value]
            if vals and set(vals) <= {'Y', 'N'}:
                col = L(cm['longterm'])
                dv = DataValidation(type='list', formula1=f"{quote_sheetname(o.title)}!${col}$1:${col}${len(vals)}", allow_blank=True)
                dv.add(f'{col}{first}:{col}{last}'); sh.add_data_validation(dv); break

    rows = []
    for r in range(first, last + 1):
        g = lambda k: sh.cell(r, cm[k]).value if cm[k] else None
        try:
            rows.append(dict(r=r, amt=float(sh.cell(r, amt_col).value), H=pdt(g('start')), I=pdt(g('end')),
                             pay=pdt(g('date')).date() if g('date') else None, acct=g('acct'),
                             close=pdt(g('close')).date() if g('close') else None,
                             ledger=g('ledger'), voucher=g('voucher'), vendor=g('vendor'), memo=g('memo')))
        except (ValueError, TypeError) as e:
            raise ScheduleError(f'{r}행의 금액 또는 날짜를 읽지 못했습니다: {e}')
    accts = list(dict.fromkeys(x['acct'] for x in rows))
    y, m = min(x['H'] for x in rows).year, min(x['H'] for x in rows).month
    ye, me = max(x['I'] for x in rows).year, max(x['I'] for x in rows).month
    months = []
    while (y, m) <= (ye, me):
        months.append(dt.date(y, m, calendar.monthrange(y, m)[1])); m += 1
        if m == 13: y, m = y + 1, 1

    has_led = bool(cm['ledger'])
    if out_path:
        _write_sheet(wb, sh, cm, S, amt_col, first, n, base, rows, accts, months, has_led, out_sheet)
        wb.save(out_path)

    # ── 독립 재계산 (Excel 수식과 동일 로직) ──
    def pybal(x, d):
        J = (x['I'] - x['H']).total_seconds() / 60
        rem = max(0.0, (x['I'] - max(dt.datetime.combine(d, dt.time()) + dt.timedelta(1), x['H'])).total_seconds() / 60)
        return r0(x['amt'] * r0(round(rem) / round(J), 5))

    NM = len(months)
    exp = [0.0] * NM; mb = [0.0] * NM; tb = 0.0; acc_b = {}; bad = []; led_bad = []
    acc_exp = {a: [0.0] * NM for a in accts}; acc_mb = {a: [0.0] * NM for a in accts}
    acc_amt = {a: 0.0 for a in accts}
    items = []
    for k, x in enumerate(rows):
        mv = [pybal(x, d.replace(day=1) - dt.timedelta(1)) - pybal(x, d) for d in months]
        mbal = [0 if (x['pay'] and d < x['pay']) else pybal(x, d) for d in months]
        if abs(sum(mv) - x['amt']) > 0.5: bad.append(x['r'])
        exp = [e + v for e, v in zip(exp, mv)]
        mb = [e + v for e, v in zip(mb, mbal)]
        acc_exp[x['acct']] = [e + v for e, v in zip(acc_exp[x['acct']], mv)]
        acc_mb[x['acct']] = [e + v for e, v in zip(acc_mb[x['acct']], mbal)]
        acc_amt[x['acct']] += x['amt']
        J = round((x['I'] - x['H']).total_seconds() / 60)
        rem = round(max(0.0, (x['I'] - max(dt.datetime.combine(base, dt.time()) + dt.timedelta(1), x['H'])).total_seconds() / 60))
        b = pybal(x, base); tb += b; acc_b[x['acct']] = acc_b.get(x['acct'], 0) + b
        led = recal = diff = None
        if has_led and x['ledger'] is not None:
            led = abs(float(x['ledger'])); recal = pybal(x, x['close'] or base); diff = recal - led
            if abs(diff) > 0.5: led_bad.append({'row': x['r'], 'no': k + 1, 'diff': diff})
        items.append(dict(
            no=k + 1, row=x['r'], voucher=_s(x['voucher']), pay=_d(x['pay']), acct=x['acct'], vendor=_s(x['vendor']),
            memo=_s(x['memo']), amt=x['amt'], start=x['H'].strftime('%Y-%m-%d %H:%M'), end=x['I'].strftime('%Y-%m-%d %H:%M'),
            total_min=J, rem_min=rem, ratio=r0(rem / J, 5) if J else 0, expense=mv, balance=mbal,
            cum=x['amt'] - b, base_bal=b, close=_d(x['close']), ledger=led, recalc=recal, diff=diff))

    warnings = []
    if cl_all and base not in set(cl_all):
        latest = max(cl_all)
        warnings.append(f'원자료에 기준일({base}) 결산일자가 없습니다. 결산일자는 {latest}까지만 있습니다. '
                        f'해당 월 원자료가 맞는지 확인하세요.')
    if base != dt.date(base.year, base.month, calendar.monthrange(base.year, base.month)[1]):
        warnings.append(f'기준일 {base}은 월말이 아닙니다. 월말 잔액 표의 기준일 월 강조가 표시되지 않습니다.')

    close_dist = {}
    for d in cl_all: close_dist[str(d)] = close_dist.get(str(d), 0) + 1

    return dict(
        sheet=sh.title, first=first, last=last, n=n, base=str(base), base_auto=base_auto,
        filename=out_filename(base), has_ledger=has_led,
        months=[d.strftime('%Y-%m') for d in months],
        total_balance=tb, total_amount=sum(x['amt'] for x in rows),
        by_account=[dict(acct=a, amount=acc_amt[a], balance=acc_b.get(a, 0), expense=acc_exp[a], month_balance=acc_mb[a]) for a in accts],
        month_expense=exp, month_balance=mb,
        check_sum_ok=not bad, check_sum_bad=bad,
        ledger_ok=(not led_bad) if has_led else None, ledger_bad=led_bad,
        close_dist=dict(sorted(close_dist.items())), warnings=warnings, items=items,
    )


def _s(v):
    return '' if v is None else str(v)


def _d(v):
    return v.isoformat() if v else ''


def _write_sheet(wb, sh, cm, S, amt_col, first, n, base, rows, accts, months, has_led, out_sheet):
    if out_sheet in wb.sheetnames: del wb[out_sheet]
    ws = wb.create_sheet(out_sheet)

    F = '맑은 고딕'
    fn = Font(name=F, size=10); fb = Font(name=F, size=10, bold=True)
    fg = Font(name=F, size=10, color='008000'); note = Font(name=F, size=9, color='808080')
    thin = Side(style='thin', color='999999'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr = PatternFill('solid', fgColor='D9E1F2'); hdr2 = PatternFill('solid', fgColor='E2EFDA'); tot = PatternFill('solid', fgColor='F2F2F2')
    NUM = '#,##0;-#,##0;"-"'
    R0 = 6; RT = R0 + n
    MC0 = 13; NM = len(months); MCL = MC0 + NM - 1
    cSum, cCum, cBal = MCL + 1, MCL + 2, MCL + 3
    cClose, cLed, cRe, cDiff = (MCL + 4, MCL + 5, MCL + 6, MCL + 7) if has_led else (None,) * 4
    cEnd = cDiff or cBal

    ws['A1'] = '선급비용 명세서'; ws['A1'].font = Font(name=F, size=16, bold=True)
    ws['A2'] = '기준일'; ws['A2'].font = fb
    ws['C2'] = base; ws['C2'].number_format = 'yyyy-mm-dd'
    ws['C2'].font = Font(name=F, size=10, bold=True, color='0000FF'); ws['C2'].fill = PatternFill('solid', fgColor='FFFF00')
    ws['D2'] = '← 기준일(월말) 입력. 기준일 24:00 시점 잔액으로 계산'; ws['D2'].font = note
    ws['A3'] = ('계산방식: 잔여비율 = 잔여기간(분) ÷ 총기간(분), 소수 5자리 반올림 / 잔액 = 발생금액 × 잔여비율(원 단위 반올림). '
                '월별 비용 = 전월말 잔액 - 당월말 잔액' + (' / 원장 대사는 각 행의 결산일자 기준 재계산' if has_led else ''))
    ws['A3'].font = note
    ws.cell(row=3, column=cEnd, value='(단위: 원)').font = fn
    ws.cell(row=3, column=cEnd).alignment = Alignment(horizontal='right')

    heads = {1: 'No', 2: '전표번호', 3: '회계일자', 4: '계정과목명', 5: '거래처명', 6: '적요', 7: '발생금액',
             8: '시작시간', 9: '종료시간', 10: '총기간(분)', 11: '기준일\n잔여기간(분)', 12: '기준일\n잔여비율',
             cSum: '월별비용 합계\n(=발생금액)', cCum: '기준일까지\n비용 누계', cBal: '기준일\n선급비용 잔액'}
    if has_led:
        heads.update({cClose: '원장\n결산일자', cLed: '원장 잔액\n(발생계산금액)', cRe: '결산일자 기준\n재계산 잔액', cDiff: '차이'})
    for c, v in heads.items(): ws.cell(row=5, column=c, value=v)
    for i in range(NM):
        c = MC0 + i
        ws.cell(row=5, column=c, value=f'=DATE({months[0].year},{months[0].month},{months[0].day})' if i == 0
                else f'=EOMONTH({L(c-1)}5,1)').number_format = 'yyyy"년" m"월"'
    ws.cell(row=4, column=MC0, value='월별 발생 비용').font = fb
    ws.cell(row=4, column=MC0 + 5, value='※ 음영: 기준일까지 비용 인식된 월').font = note
    ws.cell(row=4, column=cCum, value='기준일 현황').font = fb
    if has_led: ws.cell(row=4, column=cClose, value='원장 대사 (결산일자 기준)').font = fb
    for c in range(1, cEnd + 1):
        x = ws.cell(row=5, column=c); x.font = fb; x.fill = hdr2 if (has_led and c >= cClose) else hdr; x.border = B
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[5].height = 32

    def bal(r, cutoff):
        return f'ROUND($G{r}*ROUND(ROUND(MAX(0,$I{r}-MAX({cutoff}+1,$H{r}))*1440,0)/$J{r},5),0)'

    def ref(k, s):
        return f'{S}!{L(cm[k])}{s}' if cm[k] else None

    def dtf(rf):
        return (f'IF(ISNUMBER({rf}),{rf},DATE(LEFT({rf},4),MID({rf},5,2),MID({rf},7,2))'
                f'+TIME(IFERROR(--MID({rf},9,2),0),IFERROR(--MID({rf},11,2),0),0))')

    def df(rf):
        return f'IF(ISNUMBER({rf}),INT({rf}),DATE(LEFT({rf},4),MID({rf},5,2),MID({rf},7,2)))'

    for k in range(n):
        r = R0 + k; s = first + k
        v = {1: k + 1, 2: f'={ref("voucher", s)}' if cm['voucher'] else '', 3: '=' + df(ref('date', s)) if cm['date'] else '',
             4: f'={ref("acct", s)}', 5: f'={ref("vendor", s)}' if cm['vendor'] else '', 6: f'={ref("memo", s)}' if cm['memo'] else '',
             7: f'={S}!{L(amt_col)}{s}', 8: '=' + dtf(ref('start', s)), 9: '=' + dtf(ref('end', s)),
             10: f'=ROUND((I{r}-H{r})*1440,0)', 11: f'=ROUND(MAX(0,I{r}-MAX($C$2+1,H{r}))*1440,0)', 12: f'=ROUND(K{r}/J{r},5)'}
        for i in range(NM):
            cl = L(MC0 + i)
            v[MC0 + i] = '=' + bal(r, f'EOMONTH({cl}$5,-1)') + '-' + bal(r, f'{cl}$5')
        v[cSum] = f'=SUM({L(MC0)}{r}:{L(MCL)}{r})'
        v[cBal] = f'=ROUND(G{r}*L{r},0)'
        v[cCum] = f'=G{r}-{L(cBal)}{r}'
        if has_led:
            v[cClose] = '=' + df(ref('close', s)) if cm['close'] else '=$C$2'
            v[cLed] = f'=ABS({ref("ledger", s)})'
            v[cRe] = '=' + bal(r, f'{L(cClose)}{r}').replace('$', '')
            v[cDiff] = f'={L(cRe)}{r}-{L(cLed)}{r}'
        for c, val in v.items():
            x = ws.cell(row=r, column=c, value=val if val != '' else None); x.font = fn; x.border = B
            if c in (2, 4, 5, 6, 7) or c == cLed: x.font = fg
            if c in (3, cClose): x.number_format = 'yyyy-mm-dd'; x.alignment = Alignment(horizontal='center')
            elif c in (8, 9): x.number_format = 'yyyy-mm-dd hh:mm'; x.alignment = Alignment(horizontal='center')
            elif c in (1, 10, 11): x.number_format = '#,##0'; x.alignment = Alignment(horizontal='center')
            elif c == 12: x.number_format = '0.00000'; x.alignment = Alignment(horizontal='center')
            elif c >= 7: x.number_format = NUM
        ws.cell(row=r, column=cBal).font = Font(name=F, size=10, bold=True)

    def total_row(row, c_to):
        ws.cell(row=row, column=1, value='합계')
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
        for c in range(1, c_to + 1):
            x = ws.cell(row=row, column=c); x.font = fb; x.fill = tot; x.border = B

    total_row(RT, cEnd)
    for c in [7] + [c for c in range(MC0, cEnd + 1) if c != cClose]:
        ws.cell(row=RT, column=c, value=f'=SUM({L(c)}{R0}:{L(c)}{RT-1})').number_format = NUM
    ws.cell(row=RT, column=cBal).font = Font(name=F, size=10, bold=True, color='C00000')

    def acct_summary(top, d0, d1, c_last, title):
        ws.cell(row=top, column=1, value=title).font = fb
        cols = [7] + list(range(MC0, c_last + 1))
        for j, ac in enumerate(accts):
            rr = top + 1 + j
            ws.cell(row=rr, column=4, value=ac)
            for c in cols:
                ws.cell(row=rr, column=c, value=f'=SUMIF($D${d0}:$D${d1},$D{rr},{L(c)}${d0}:{L(c)}${d1})').number_format = NUM
            for c in range(1, c_last + 1):
                x = ws.cell(row=rr, column=c); x.font = fn; x.border = B
        rr = top + 1 + len(accts)
        ws.cell(row=rr, column=4, value='합계')
        for c in cols:
            ws.cell(row=rr, column=c, value=f'=SUM({L(c)}{top+1}:{L(c)}{rr-1})').number_format = NUM
        for c in range(1, c_last + 1):
            x = ws.cell(row=rr, column=c); x.font = fb; x.fill = tot; x.border = B
        return rr

    rs_end = acct_summary(RT + 2, R0, RT - 1, cBal, '[계정과목별 요약]')
    ws.cell(row=rs_end, column=cBal).font = Font(name=F, size=10, bold=True, color='C00000')
    ws.conditional_formatting.add(f'{L(MC0)}{R0}:{L(MCL)}{rs_end}',
        FormulaRule(formula=[f'{L(MC0)}$5<=$C$2'], fill=PatternFill('solid', fgColor='FFF2CC')))
    if has_led:
        ws.conditional_formatting.add(f'{L(cDiff)}{R0}:{L(cDiff)}{RT}',
            FormulaRule(formula=[f'{L(cDiff)}{R0}<>0'], font=Font(color='FF0000', bold=True)))

    RB = rs_end + 3; HB = RB + 1; D0 = HB + 1; DT = D0 + n
    ws.cell(row=RB, column=1, value='[월별 선급비용 잔액]').font = Font(name=F, size=12, bold=True)
    ws.cell(row=RB, column=MC0, value='각 월말 24:00 기준 잔액 (전월말 잔액 + 당월 신규 발생(회계일자 기준) - 당월 비용 = 당월말 잔액). 회계일자 이전 월은 0').font = note
    ws.cell(row=RB, column=MC0 + 8, value='※ 붉은 음영: 기준일 월').font = note
    for c, val in {1: 'No', 2: '전표번호', 3: '회계일자', 4: '계정과목명', 5: '거래처명', 6: '적요', 7: '발생금액',
                   8: '시작시간', 9: '종료시간', cSum: '검증\n(기준일 잔액 차이)'}.items():
        ws.cell(row=HB, column=c, value=val)
    for i in range(NM):
        ws.cell(row=HB, column=MC0 + i, value=f'={L(MC0+i)}$5').number_format = 'yyyy"년" m"월"'
    for c in list(range(1, MCL + 1)) + [cSum]:
        x = ws.cell(row=HB, column=c); x.font = fb; x.fill = hdr; x.border = B
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[HB].height = 32
    for k in range(n):
        rr = D0 + k; rm = R0 + k
        v = {1: k + 1, **{c: f'={L(c)}{rm}' for c in range(2, 10)}}
        for i in range(NM):
            cl = L(MC0 + i); b = bal(rm, f'{cl}${HB}')
            v[MC0 + i] = f'=IF(AND(ISNUMBER($C{rm}),{cl}${HB}<$C{rm}),0,{b})'
        v[cSum] = f'=IFERROR(INDEX({L(MC0)}{rr}:{L(MCL)}{rr},MATCH($C$2,{L(MC0)}${HB}:{L(MCL)}${HB},0))-{L(cBal)}{rm},"-")'
        for c, val in v.items():
            x = ws.cell(row=rr, column=c, value=val); x.font = fn; x.border = B
            if c == 3: x.number_format = 'yyyy-mm-dd'; x.alignment = Alignment(horizontal='center')
            elif c in (8, 9): x.number_format = 'yyyy-mm-dd hh:mm'; x.alignment = Alignment(horizontal='center')
            elif c == 1: x.alignment = Alignment(horizontal='center')
            elif c >= 7: x.number_format = NUM
        ws.cell(row=rr, column=cSum).alignment = Alignment(horizontal='center')
    total_row(DT, MCL)
    for c in [7] + list(range(MC0, MCL + 1)) + [cSum]:
        x = ws.cell(row=DT, column=c, value=f'=SUM({L(c)}{D0}:{L(c)}{DT-1})'); x.number_format = NUM; x.font = fb; x.fill = tot; x.border = B
    rb_end = acct_summary(DT + 2, D0, DT - 1, MCL, '[계정과목별 월별 잔액 요약]')
    ws.conditional_formatting.add(f'{L(MC0)}{HB}:{L(MCL)}{rb_end}',
        FormulaRule(formula=[f'{L(MC0)}${HB}=$C$2'], fill=PatternFill('solid', fgColor='FCE4D6'), font=Font(bold=True, color='C00000')))
    ws.conditional_formatting.add(f'{L(MC0)}{D0}:{L(MCL)}{rb_end}',
        FormulaRule(formula=[f'{L(MC0)}${HB}<$C$2'], fill=PatternFill('solid', fgColor='FFF2CC')))
    ws.conditional_formatting.add(f'{L(cSum)}{D0}:{L(cSum)}{DT}',
        FormulaRule(formula=[f'AND(ISNUMBER({L(cSum)}{D0}),{L(cSum)}{D0}<>0)'], font=Font(color='FF0000', bold=True)))

    for c, w in {1: 5, 2: 15, 3: 11, 4: 22, 5: 18, 6: 46, 7: 13, 8: 16, 9: 16, 10: 10, 11: 11, 12: 9}.items():
        ws.column_dimensions[L(c)].width = w
    for c in range(MC0, cEnd + 1): ws.column_dimensions[L(c)].width = 13
    if cClose: ws.column_dimensions[L(cClose)].width = 11
    ws.freeze_panes = f'{L(MC0)}{R0}'
    ws.sheet_view.zoomScale = 90
    wb.active = wb.sheetnames.index(out_sheet)
    wb.calculation.fullCalcOnLoad = True
