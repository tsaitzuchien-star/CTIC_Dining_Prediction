# -*- coding: utf-8 -*-
"""
中創園區天氣資料（2026-11-13）

院內網路連不到天氣網站，所以由 GitHub Actions 每天去 Open-Meteo 抓，存成 data/weather.csv，
筆電（ctic_daily_record.py）、戰情室網站、訂餐 App 都從這個檔讀。

每列一天：日期、07~12 點每小時降雨量(mm)、09 點天氣代碼、07~12 點累積雨量、09 點是否下雨。
是否下雨的判斷與原本相同：該小時降雨量 > 0 或天氣代碼 >= 51（毛毛雨以上）。
"""
import csv
import json
import os
import urllib.request
from datetime import date, timedelta

LAT, LON = 23.9, 120.68
START = date(2026, 3, 30)          # 卡機資料最早的日期
HOURS = range(7, 13)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'weather.csv')
FIELDS = ['日期'] + [f'雨量{h:02d}' for h in HOURS] + ['代碼09', '上午雨量', '09點下雨']


def fetch(url):
    with urllib.request.urlopen(url, timeout=60) as r:
        return json.load(r)


def hourly_rows(d):
    rows = {}
    for t, p, c in zip(d['hourly']['time'], d['hourly']['precipitation'], d['hourly']['weather_code']):
        day, hh = t[:10], int(t[11:13])
        rows.setdefault(day, {})[hh] = (p or 0.0, c or 0)
    return rows


def main():
    common = f'latitude={LAT}&longitude={LON}&hourly=precipitation,weather_code&timezone=Asia%2FTaipei'
    data = {}
    # 歷史：archive 有約 5 天延遲
    end_hist = date.today() - timedelta(days=6)
    if end_hist >= START:
        data.update(hourly_rows(fetch(f'https://archive-api.open-meteo.com/v1/archive?{common}'
                                      f'&start_date={START}&end_date={end_hist}')))
    # 最近幾天＋今天：forecast 的 past_days（今天的值會隨觀測更新）
    data.update(hourly_rows(fetch(f'https://api.open-meteo.com/v1/forecast?{common}&past_days=7&forecast_days=1')))

    out = []
    for day in sorted(data):
        h = data[day]
        if not all(x in h for x in HOURS):
            continue
        p9, c9 = h[9]
        out.append({'日期': day, **{f'雨量{x:02d}': round(h[x][0], 1) for x in HOURS}, '代碼09': c9,
                    '上午雨量': round(sum(h[x][0] for x in HOURS), 1),
                    '09點下雨': '是' if (p9 > 0 or c9 >= 51) else '否'})
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(out)
    print(f'{len(out)} 天，{out[0]["日期"]} ~ {out[-1]["日期"]}，09 點下雨 {sum(r["09點下雨"] == "是" for r in out)} 天')


if __name__ == '__main__':
    main()
