# -*- coding: utf-8 -*-
"""
GitHub Actions で毎晩実行：B3 の株式（普通株・優先株・ユニット）の株価をヤフーから取り、
Web ページ用のデータ（site/data/all.json）を作る。
MA の角度や絞り込みはページ側（ブラウザ）で計算するので、ここでは終値の並びと基本の数値だけを書き出す。
"""
import datetime as dt
import json
import math
import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import screener_core as core   # パソコン版と同じデータ取得・計算の部品を使う

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "site", "data")
MIN_OK = int(os.environ.get("MIN_OK", "20"))   # 取得できた銘柄がこれより少なければ失敗扱い
BARS = {"d": 300, "w": 200, "m": 120}     # ページに渡す本数（日足・週足・月足）
core.HISTORY_PERIOD = "10y"               # 月足の長い MA にも対応できるように 10 年分


def r2(x):
    try:
        x = float(x)
        return None if math.isnan(x) or math.isinf(x) else round(x, 2)
    except Exception:
        return None


def main():
    codes, tinfo = core.load_tickers("全銘柄")
    print(f"対象 {len(codes)} 銘柄")
    prices = core.fetch_prices(codes)
    print(f"株価を取得できた銘柄: {len(prices)}")
    if len(prices) < MIN_OK:
        raise SystemExit("株価をほとんど取得できませんでした（ヤフーに断られた可能性）")

    series = {}
    for c, df in prices.items():
        series[c] = {
            "d": df["Close"],
            "w": core.resample_close(df, "W-FRI"),
            "m": core.resample_close(df, "ME"),
        }
    # 足ごとに共通の日付の並びを作り、各銘柄の終値をそろえる（取引のない日は直前の値）
    dates, aligned = {}, {c: {} for c in series}
    for tf, n in BARS.items():
        idx = sorted(set().union(*[set(s[tf].index) for s in series.values()]))[-n:]
        dates[tf] = [d.strftime("%Y-%m-%d") for d in idx]
        for c, s in series.items():
            v = s[tf].reindex(idx).ffill(limit=5)
            aligned[c][tf] = [r2(x) for x in v.values]

    stocks = []
    for c, df in prices.items():
        cl, v = df["Close"], df["Volume"]
        last = cl.iloc[-1]
        prev = cl.iloc[-2] if len(cl) > 1 else np.nan
        tr = pd.concat([df["High"] - df["Low"], (df["High"] - cl.shift()).abs(), (df["Low"] - cl.shift()).abs()], axis=1).max(axis=1)
        info = tinfo.get(c, {})
        stocks.append({
            "c": c,
            "n": info.get("name", ""),
            "s": info.get("sector", ""),
            "p": r2(last),
            "chg": r2((last / prev - 1) * 100) if prev else None,
            "val": r2((cl * v).tail(20).mean() / 1e6),          # 20日平均売買代金（百万R$）
            "atr": r2(tr.tail(14).mean() / last * 100) if last else None,
            "rsi": r2(core.rsi(cl).iloc[-1]),
            "last": df.index[-1].strftime("%Y-%m-%d"),
            "d": aligned[c]["d"], "w": aligned[c]["w"], "m": aligned[c]["m"],
        })
    stocks.sort(key=lambda s: -(s["val"] or 0))

    os.makedirs(OUT_DIR, exist_ok=True)
    now = dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))
    out = {"updated": now.strftime("%Y-%m-%d %H:%M") + "（日本時間）", "dates": dates, "stocks": stocks}
    path = os.path.join(OUT_DIR, "all.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"書き出し: {path}（{os.path.getsize(path) / 1e6:.1f} MB、{len(stocks)} 銘柄）")


if __name__ == "__main__":
    main()
