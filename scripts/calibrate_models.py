# -*- coding: utf-8 -*-
"""
中創園區餐廳預測：全自動校準與選模（2026-11-18）

GitHub Actions 每個上班日下午（一條龍寫完趨勢紀錄後）執行，不用任何人動手：

1. 讀 Google Sheets 已發布的 CSV（趨勢紀錄、團膳訂單、預測快照）＋ data/weather.csv ＋ 國定假日。
2. 對每個上班日，只用「那天以前」的資料，重算各種方法當天會給的預測（逐日回測，不偷看答案），
   跟實際用餐（刷卡＋現金）比較，寫到 data/model_daily.csv。
3. 依最近的誤差自動決定：
   - 08:30 第一版用哪個方法（近 8 次同星期中位數 hist8，或混入 08:20 已到常客的 blend0820）
     ＋混合比重；挑戰者近 N 個工作日平均誤差比現行方法少 SWITCH_MARGIN 份以上才切換，反過來也會自動切回。
   - 08:20、09:20 出席率（用快照當下的人數算，跟早上實際看到的人數同一個口徑）
   - 連假前一天係數（過去連假前一天 實際÷第一版 的中位數，滿 3 次才啟用，而且要真的有讓誤差變小）
   - App「建議份數」裡進園推估的比重（原本固定各半）
   結果寫到 data/model_params.json，戰情室、訂餐 App 每次開啟都讀這個檔。
4. 每週一（或有切換時）產生一則報告到 data/model_report.json，筆電 model_notify.py 讀到新報告就 LINE 給 Cliff。

本機測試：CAL_LOCAL_DIR=資料夾（放 trend.csv、order.csv、snapshot.csv）python scripts/calibrate_models.py
"""
import csv
import io
import json
import os
import statistics as st
import urllib.request
from datetime import date, datetime, timedelta, timezone

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
DATA = os.path.join(ROOT, 'data')
PUB = ('https://docs.google.com/spreadsheets/d/e/2PACX-1vSv8X_NJDPjSVJTEkFjdBTSJKjDnZ18eyf6yBdTBq2eVT8TB6_wB-'
       'UDUUVUoF5eUJrLSnDsqpu7OhbS/pub?single=true&output=csv&gid=')
SOURCES = {'trend': PUB + '1825976142', 'order': PUB + '830830', 'snapshot': PUB + '397081291'}
HOLIDAY_URL = 'https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/{y}.json'
TPE = timezone(timedelta(hours=8))

HIST_LOOKBACK = 8          # 第一版：近 8 次同星期
RATIO_DAYS, RATIO_FRI, CALIB_MIN = 10, 6, 3
FRIDAY_FACTOR = 0.85
WEIGHTS = [0, 0.25, 0.5, 0.75, 1]
FIT_DAYS = 20              # 挑比重看最近幾個工作日
EVAL_DAYS = 10             # 切換方法看最近幾個工作日
MIN_EVAL = 8               # 至少要有幾天可比較才會切換
SWITCH_MARGIN = 1.0        # 挑戰者平均誤差要少幾份才切換
DEFAULT_DOOR_WEIGHT = 0.5  # App 建議份數原本是歷史與進園各半

# 2026 平日放假日（行政院人事行政總處公告）；抓得到網路上的行事曆時以網路為準
HOLIDAYS_2026 = ['2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20',
                 '2026-02-27', '2026-04-03', '2026-04-06', '2026-05-01', '2026-06-19', '2026-09-25',
                 '2026-09-28', '2026-10-09', '2026-10-26', '2026-12-25']


# ---------- 讀資料 ----------
def read_csv(name):
    local = os.environ.get('CAL_LOCAL_DIR')
    try:
        if local:
            with open(os.path.join(local, f'{name}.csv'), encoding='utf-8') as f:
                text = f.read()
        else:
            with urllib.request.urlopen(SOURCES[name], timeout=60) as r:
                text = r.read().decode('utf-8')
        rows = list(csv.DictReader(io.StringIO(text)))
        if rows and '<html' in text[:200].lower():
            raise ValueError('不是 CSV（可能沒有發布）')
        return rows, None
    except Exception as e:
        return [], f'{type(e).__name__}: {e}'


def to_int(v):
    try:
        return int(float(str(v).strip()))
    except (TypeError, ValueError):
        return None


def norm_date(s):
    s = str(s or '').strip().replace('/', '-')
    try:
        y, m, d = s[:10].split('-')[:3]
        return date(int(y), int(m), int(d[:2]))
    except Exception:
        return None


def load_holidays(years):
    """回傳 (平日放假日 set, 補班日 set, 來源說明)"""
    off, work, src = set(), set(), []
    for y in years:
        cache = os.path.join(DATA, f'holidays_{y}.json')
        rows = None
        if not os.environ.get('CAL_LOCAL_DIR'):
            try:
                with urllib.request.urlopen(HOLIDAY_URL.format(y=y), timeout=30) as r:
                    rows = json.load(r)
                with open(cache, 'w', encoding='utf-8') as f:
                    json.dump(rows, f, ensure_ascii=False)
            except Exception:
                rows = None
        if rows is None and os.path.exists(cache):
            with open(cache, encoding='utf-8') as f:
                rows = json.load(f)
        if rows:
            for r in rows:
                d = datetime.strptime(r['date'], '%Y%m%d').date()
                if r.get('isHoliday') and d.weekday() < 5:
                    off.add(d)
                elif not r.get('isHoliday') and d.weekday() >= 5:
                    work.add(d)
            src.append(f'{y} 政府行事曆')
        elif y == 2026:
            off |= {norm_date(x) for x in HOLIDAYS_2026}
            src.append('2026 內建清單')
        else:
            src.append(f'{y} 無資料')
    return off, work, '、'.join(src)


def is_workday(d, off, work):
    return d in work or (d.weekday() < 5 and d not in off)


def is_holiday_eve(d, off, work):
    """連假前一天：下一個上班日在 4 天以上之後（例：週四接週五放假、週五接週一放假），或平日隔天就放假"""
    if not is_workday(d, off, work):
        return False
    n = d + timedelta(days=1)
    while not is_workday(n, off, work):
        n += timedelta(days=1)
    gap = (n - d).days
    return gap >= 4 or (gap >= 2 and d.weekday() < 4)


def build_days(trend, order, snap, wx):
    order_by = {norm_date(r.get('日期')): r for r in order}
    snap_by = {}
    for r in snap:
        d, slot, n = norm_date(r.get('日期')), str(r.get('時段', '')).strip().zfill(5), to_int(r.get('已到常客'))
        if d and n:
            snap_by.setdefault(d, {})[slot] = n
    days = {}
    for r in trend:
        d, act = norm_date(r.get('日期')), to_int(r.get('實際用餐'))
        if not d or not act or d.weekday() >= 5:
            continue
        o = order_by.get(d, {})
        cash = to_int(r.get('現金付款'))
        if cash is None:
            cash = to_int(o.get('現場付現'))
        s = snap_by.get(d, {})
        days[d] = {'date': d, 'card': act, 'cash': cash, 'order': to_int(o.get('用餐人數')),
                   'meat': str(o.get('主菜類型', '')).strip(), 'reason': str(o.get('加量原因', '')).strip(),
                   'add': to_int(o.get('臨時加量')), 'missed': to_int(o.get('沒吃到人數')),
                   'n0820': s.get('08:20'), 'n0920': s.get('09:20'),
                   'n0820_eod': to_int(r.get('08:20已到常客')), 'n0920_eod': to_int(r.get('09:20已到常客')),
                   'rain': (wx.get(d.isoformat()) or {}).get('09點下雨') == '是'}
    # 現金沒填的日子（6/15 開始記現金）用最近 10 個有填的中位數補，6/15 以前當 0（與戰情室口徑相同）
    known = []
    for d in sorted(days):
        x = days[d]
        if x['cash'] is None:
            x['cash_filled'] = True
            x['cash'] = int(round(st.median(known[-10:]))) if known and d >= date(2026, 6, 15) else 0
        else:
            known.append(x['cash'])
        x['actual'] = x['card'] + x['cash']
    return [days[d] for d in sorted(days)]


# ---------- 各方法（只用 i 之前的資料）----------
def med(v):
    return st.median(v) if v else None


def hist8(days, i):
    d = days[i]['date']
    same = [x['actual'] for x in days[:i] if x['date'].weekday() == d.weekday()][-HIST_LOOKBACK:]
    return round(st.median(same)) if len(same) >= 4 else None


def ratio(days, i, key):
    """出席率：實際 ÷ 當時已到常客。平日＝最近 10 個非週五、沒下雨；週五＝最近 6 個週五（不足 3 個用平日×0.85）"""
    fri = days[i]['date'].weekday() == 4
    prev = [x for x in days[:i] if x.get(key) and not x['rain']]
    base = [x['actual'] / x[key] for x in prev if x['date'].weekday() != 4][-RATIO_DAYS:]
    r = med(base) if len(base) >= CALIB_MIN else None
    if fri:
        f = [x['actual'] / x[key] for x in prev if x['date'].weekday() == 4][-RATIO_FRI:]
        if len(f) >= CALIB_MIN:
            return med(f), len(f)
        return (r * FRIDAY_FACTOR, 0) if r else (None, 0)
    return r, len(base)


def compute(days, off, work):
    rows = []
    for i, x in enumerate(days):
        h = hist8(days, i)
        r8, _ = ratio(days, i, 'n0820')
        r9, _ = ratio(days, i, 'n0920')
        reg8 = round(x['n0820'] * r8) if x['n0820'] and r8 else None
        door9 = round(x['n0920'] * r9) if x['n0920'] and r9 else None
        rows.append({'date': x['date'], 'actual': x['actual'], 'hist8': h, 'reg0820': reg8, 'door0920': door9,
                     'order': x['order'], 'eve': is_holiday_eve(x['date'], off, work), 'x': x})
    # 連假前一天係數：過去連假前一天 實際÷hist8 的中位數（逐日只用以前的）
    for i, r in enumerate(rows):
        prev = [p['actual'] / p['hist8'] for p in rows[:i] if p['eve'] and p['hist8']]
        r['eve_factor'] = med(prev) if len(prev) >= CALIB_MIN else None
        r['eve_n'] = len(prev)
    return rows


def mae(rows, key):
    e = [abs(r[key] - r['actual']) for r in rows if r.get(key) is not None]
    return (round(sum(e) / len(e), 1), len(e)) if e else (None, 0)


def blend(a, b, w):
    if a is None:
        return None
    if b is None:
        return a if w < 1 else None
    return round((1 - w) * a + w * b)


def fit_weight(rows, base_key, door_key, default):
    """在最近 FIT_DAYS 個兩者都有的工作日裡，挑平均誤差最小的比重（同分取較接近 default 的）"""
    ok = [r for r in rows if r.get(base_key) is not None and r.get(door_key) is not None][-FIT_DAYS:]
    if len(ok) < MIN_EVAL:
        return default, None, len(ok)
    best = None
    for w in WEIGHTS:
        e = sum(abs(blend(r[base_key], r[door_key], w) - r['actual']) for r in ok) / len(ok)
        key = (round(e, 2), abs(w - default))
        if best is None or key < best[0]:
            best = (key, w, round(e, 1))
    return best[1], best[2], len(ok)


def walk_forward(rows):
    """逐日：用前面的資料挑比重，再產生當天 blend0820、加連假係數、App 建議的預測"""
    for i, r in enumerate(rows):
        w8, _, _ = fit_weight(rows[:i], 'hist8', 'reg0820', 0.0)
        r['w0820'] = w8
        r['blend0820'] = blend(r['hist8'], r['reg0820'], w8) if r['reg0820'] is not None else None
        for k in ('hist8', 'blend0820'):
            v = r[k]
            r[k + '_eve'] = round(v * r['eve_factor']) if v is not None and r['eve'] and r['eve_factor'] else v
        r['suggest_half'] = blend(r['hist8'], r['door0920'], 0.5) if r['door0920'] is not None else None
    for i, r in enumerate(rows):
        wd, _, _ = fit_weight(rows[:i], 'hist8', 'door0920', DEFAULT_DOOR_WEIGHT)
        r['wdoor'] = wd
        r['suggest_auto'] = blend(r['hist8'], r['door0920'], wd) if r['door0920'] is not None else None


# ---------- 決定今天起要用的參數 ----------
def decide(rows, prev_params, days, today):
    out = {}
    last = rows[-EVAL_DAYS:]
    incumbent = (prev_params.get('first_version') or {}).get('method', 'hist8')
    cands = {}
    for m in ('hist8', 'blend0820'):
        both = [r for r in rows if r.get('hist8') is not None and r.get('blend0820') is not None][-EVAL_DAYS:]
        cands[m] = mae(both, m)
    chosen, why = incumbent, '維持'
    a, b = cands['hist8'], cands['blend0820']
    if a[1] >= MIN_EVAL and b[1] >= MIN_EVAL:
        inc, ch = ('hist8', 'blend0820') if incumbent == 'hist8' else ('blend0820', 'hist8')
        if cands[ch][0] <= cands[inc][0] - SWITCH_MARGIN:
            chosen, why = ch, f'近 {cands[ch][1]} 個工作日平均誤差 {cands[ch][0]} 份，比 {inc} 的 {cands[inc][0]} 份少 {SWITCH_MARGIN} 份以上'
    else:
        why = f'08:20 快照可比較天數 {b[1]} 天，未滿 {MIN_EVAL} 天，先用 hist8'
        chosen = 'hist8' if b[1] < MIN_EVAL else chosen

    # 現在這一刻要給前端用的係數（用全部資料）
    n = len(days)
    probe = days + [{'date': today, 'rain': False, 'actual': 0}]
    r8, n8 = ratio(probe, n, 'n0820')
    r9, n9 = ratio(probe, n, 'n0920')
    fr = dict(probe[-1], date=today + timedelta(days=(4 - today.weekday()) % 7 or 7) if today.weekday() != 4 else today)
    probe_f = days + [fr]
    fr8, fn8 = ratio(probe_f, n, 'n0820')
    fr9, fn9 = ratio(probe_f, n, 'n0920')
    w8, w8_mae, w8_n = fit_weight(rows, 'hist8', 'reg0820', 0.0)
    wd, wd_mae, wd_n = fit_weight(rows, 'hist8', 'door0920', DEFAULT_DOOR_WEIGHT)

    eve_pts = [(r['date'], r['actual'], r['hist8']) for r in rows if r['eve'] and r['hist8']]
    eve_factor = med([a / h for _, a, h in eve_pts]) if len(eve_pts) >= CALIB_MIN else None
    eve_on = False
    if eve_factor:
        with_f = [r for r in rows if r['eve'] and r['hist8_eve'] is not None and r['eve_factor']]
        if len(with_f) >= 1:
            eve_on = mae(with_f, 'hist8_eve')[0] <= mae(with_f, 'hist8')[0]
        else:
            eve_on = abs(1 - eve_factor) >= 0.05   # 還沒有可驗證的次數時，差 5% 以上才用

    out['first_version'] = {
        'method': chosen, 'reason': why, 'w0820': w8, 'w0820_fit_mae': w8_mae, 'w0820_fit_days': w8_n,
        'mae_recent': {k: {'mae': v[0], 'days': v[1]} for k, v in cands.items()}}
    if chosen != incumbent:
        out['first_version']['switched_from'] = incumbent
        out['first_version']['switched_on'] = today.isoformat()
    out['ratios'] = {
        'r0820': round(r8, 3) if r8 else None, 'r0820_days': n8,
        'r0920': round(r9, 3) if r9 else None, 'r0920_days': n9,
        'fr0820': round(fr8, 3) if fr8 else None, 'fr0820_days': fn8,
        'fr0920': round(fr9, 3) if fr9 else None, 'fr0920_days': fn9,
        'note': '出席率＝實際用餐（含現金）÷ 預測快照當下的已到常客（與早上看到的人數同口徑）'}
    out['suggest'] = {'door_weight': wd, 'fit_mae': wd_mae, 'fit_days': wd_n,
                      'note': 'App 建議份數 = 第一版 ×(1−比重) + 進園推估 × 比重'}
    out['holiday_eve'] = {'factor': round(eve_factor, 3) if eve_factor else None, 'enabled': bool(eve_on and eve_factor),
                          'events': [{'date': d.isoformat(), 'actual': a, 'hist8': h} for d, a, h in eve_pts]}
    return out


def main():
    today = datetime.now(TPE).date()
    errors = {}
    trend, errors['trend'] = read_csv('trend')
    order, errors['order'] = read_csv('order')
    snap, errors['snapshot'] = read_csv('snapshot')
    if not trend:
        raise SystemExit(f'讀不到趨勢紀錄：{errors["trend"]}')
    wx = {}
    wpath = os.path.join(DATA, 'weather.csv')
    if os.path.exists(wpath):
        with open(wpath, encoding='utf-8') as f:
            wx = {r['日期']: r for r in csv.DictReader(f)}
    days = build_days(trend, order, snap, wx)
    off, work, hsrc = load_holidays(sorted({days[0]['date'].year, today.year, today.year + 1}))
    rows = compute(days, off, work)
    walk_forward(rows)

    ppath = os.path.join(DATA, 'model_params.json')
    prev = {}
    if os.path.exists(ppath):
        with open(ppath, encoding='utf-8') as f:
            prev = json.load(f)
    params = decide(rows, prev, days, today)
    upcoming = []
    d = today
    while len(upcoming) < 3 and d < today + timedelta(days=120):
        if is_holiday_eve(d, off, work):
            upcoming.append(d.isoformat())
        d += timedelta(days=1)
    params['holiday_eve']['upcoming'] = upcoming
    params['holiday_eve']['calendar'] = hsrc
    params['updated'] = datetime.now(TPE).strftime('%Y-%m-%d %H:%M')
    params['data_until'] = days[-1]['date'].isoformat()
    params['sources'] = {k: ('ok' if not v else v) for k, v in errors.items()}
    if errors['snapshot']:
        params['first_version']['method'] = 'hist8'
        params['first_version']['reason'] = '讀不到預測快照，先用 hist8'
    params['history'] = (prev.get('history') or [])[-30:]
    if params['first_version'].get('switched_from'):
        params['history'].append({'date': today.isoformat(), 'from': params['first_version']['switched_from'],
                                  'to': params['first_version']['method'], 'reason': params['first_version']['reason']})

    # 每日紀錄
    keys = ['hist8', 'reg0820', 'w0820', 'blend0820', 'hist8_eve', 'blend0820_eve', 'door0920', 'wdoor',
            'suggest_half', 'suggest_auto', 'order']
    with open(os.path.join(DATA, 'model_daily.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['日期', '實際(含現金)', '現金為補值', '連假前一天', '08:20快照常客', '09:20快照常客', '下雨'] + keys)
        for r in rows:
            x = r['x']
            w.writerow([r['date'], r['actual'], '是' if x.get('cash_filled') else '', '是' if r['eve'] else '',
                        x['n0820'] or '', x['n0920'] or '', '是' if x['rain'] else ''] + ['' if r.get(k) is None else r[k] for k in keys])

    # 成績（最近 10、20 個工作日；全部）
    keys_b = ('hist8', 'reg0820', 'blend0820', 'hist8_eve', 'door0920', 'suggest_half', 'suggest_auto', 'order')

    def board(n, common=False):
        sub = [r for r in rows if all(r.get(k) is not None for k in keys_b)] if common else rows
        sub = sub[-n:] if n else sub
        return {k: dict(zip(('mae', 'days'), mae(sub, k))) for k in keys_b}
    # common = 各方法都有預測的同一批日子（比較才公平）；last10/last20/all = 各方法各自有資料的日子
    params['scoreboard'] = {'common10': board(10, True), 'common_all': board(0, True),
                            'last10': board(10), 'last20': board(20), 'all': board(0)}
    params['factors_extra'] = extra_stats(rows)
    with open(ppath, 'w', encoding='utf-8') as f:
        json.dump(params, f, ensure_ascii=False, indent=2, default=str)
    write_report(params, today, prev)
    fv = params['first_version']
    print(f"資料到 {params['data_until']}；第一版 {fv['method']}（{fv['reason']}）；08:20 比重 {fv['w0820']}；"
          f"進園比重 {params['suggest']['door_weight']}；連假係數 {params['holiday_eve']}")
    print(json.dumps(params['scoreboard'], ensure_ascii=False))


def extra_stats(rows):
    """散肉日、臨時加減量原因：先統計，給報告參考（影響的是加購主菜份數，不是用餐人數）"""
    meat = [r['x'] for r in rows if r['x']['meat']]
    out = {}
    for kind in ('散肉', '整塊'):
        g = [x for x in meat if x['meat'] == kind]
        miss = [x['missed'] for x in g if x['missed'] is not None]
        res = [x['actual'] - x['order'] for x in g if x['order']]
        out[kind] = {'days': len(g), 'avg_missed': round(sum(miss) / len(miss), 1) if miss else None,
                     'avg_actual_minus_order': round(sum(res) / len(res), 1) if res else None}
    reasons = {}
    for r in rows:
        x = r['x']
        if x['add'] is None or not x['order']:
            continue
        k = x['reason'] or '（未填）'
        reasons.setdefault(k, []).append((x['add'], x['actual'] - x['order']))
    out['reasons'] = {k: {'days': len(v), 'avg_add': round(sum(a for a, _ in v) / len(v), 1),
                          'avg_actual_minus_order': round(sum(b for _, b in v) / len(v), 1)} for k, v in reasons.items()}
    return out


NAMES = {'hist8': '近8次同星期中位數', 'reg0820': '08:20已到常客×出席率', 'blend0820': '中位數＋08:20常客混合',
         'hist8_eve': '中位數＋連假前減量', 'door0920': '09:20進園推估', 'suggest_half': 'App建議（各半，舊）',
         'suggest_auto': 'App建議（自動比重）', 'order': '實際訂量'}


def write_report(p, today, prev):
    """每週一，或第一版方法有切換時，產生一則新報告給 LINE"""
    fv = p['first_version']
    switched = bool(fv.get('switched_from'))
    rpath = os.path.join(DATA, 'model_report.json')
    old = {}
    if os.path.exists(rpath):
        with open(rpath, encoding='utf-8') as f:
            old = json.load(f)
    weekly_due = today.weekday() == 0 or not old
    if not (switched or weekly_due) or old.get('id') == today.isoformat():
        return
    sb = p['scoreboard']['common10']
    lines = [f"🤖 預測模型{'切換通知' if switched else '週報'}（資料到 {p['data_until']}）"]
    if switched:
        lines.append(f"第一版改用：{NAMES.get(fv['method'], fv['method'])}（原 {NAMES.get(fv['switched_from'])}）")
        lines.append(f"原因：{fv['reason']}")
    else:
        lines.append(f"第一版：{NAMES.get(fv['method'], fv['method'])}（{fv['reason']}）")
    n_common = max((v.get('days') or 0) for v in sb.values())
    lines.append(f'同一批 {n_common} 個工作日的平均誤差（份）：')
    for k in ('hist8', 'blend0820', 'door0920', 'suggest_half', 'suggest_auto', 'order'):
        v = sb.get(k) or {}
        if v.get('mae') is not None:
            lines.append(f"・{NAMES[k]} {v['mae']}")
    r = p['ratios']
    lines.append(f"出席率 08:20×{r['r0820']}、09:20×{r['r0920']}；App 進園比重 {p['suggest']['door_weight']}")
    he = p['holiday_eve']
    if he['upcoming']:
        lines.append(f"連假前一天 {he['upcoming'][0]}：" + (f"自動×{he['factor']}" if he['enabled'] else '資料不足，暫不減量'))
    lines.append('不用動手，系統每天自動校準。')
    rep = {'id': today.isoformat(), 'kind': 'switch' if switched else 'weekly', 'text': '\n'.join(lines)}
    with open(rpath, 'w', encoding='utf-8') as f:
        json.dump(rep, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    main()
