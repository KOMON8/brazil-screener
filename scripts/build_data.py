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


FUND_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fund_cache.json")
FUND_DAYS = 7                      # 財務の数字はこの日数ごとに取り直す（決算は3か月ごとなので週1回で十分）


def fetch_fundamentals(codes):
    """ROE などの財務の数字（ヤフーの銘柄情報）。前回の記録が新しい銘柄は取り直さない"""
    import time
    import yfinance as yf
    today = dt.date.today()
    cache = {}
    if os.path.exists(FUND_CACHE):
        try:
            with open(FUND_CACHE, encoding="utf-8") as f:
                cache = json.load(f)
        except Exception:
            cache = {}
    todo = [c for c in codes
            if c not in cache or "roa" not in cache[c] or (today - dt.date.fromisoformat(cache[c].get("checked", "2000-01-01"))).days >= FUND_DAYS]
    print(f"財務の数字を取得中（{len(todo)} 銘柄、残りは記録を使う）...")
    for i, c in enumerate(todo):
        rec = {"checked": today.isoformat()}
        try:
            info = yf.Ticker(c + ".SA").info
            roe = info.get("returnOnEquity")
            rec["roe"] = r2(roe * 100) if isinstance(roe, (int, float)) else None
            roa = info.get("returnOnAssets")
            rec["roa"] = r2(roa * 100) if isinstance(roa, (int, float)) else None
            rec["sector"] = info.get("sector") or ""
        except Exception as e:
            rec["err"] = str(e)[:80]
            rec["checked"] = (today - dt.timedelta(days=FUND_DAYS)).isoformat()   # 失敗したら次回また取る
        cache[c] = rec
        time.sleep(0.4)
        if (i + 1) % 100 == 0:
            print(f"  {i + 1} / {len(todo)}")
    try:
        with open(FUND_CACHE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass
    return cache


def main():
    codes, tinfo = core.load_tickers("全銘柄")
    print(f"対象 {len(codes)} 銘柄")
    prices = core.fetch_prices(codes)
    print(f"株価を取得できた銘柄: {len(prices)}")
    if len(prices) < MIN_OK:
        raise SystemExit("株価をほとんど取得できませんでした（ヤフーに断られた可能性）")

    fund = fetch_fundamentals(list(prices))

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
            "roe": (fund.get(c) or {}).get("roe"),
            "roa": (fund.get(c) or {}).get("roa"),
            "fin": ((fund.get(c) or {}).get("sector") or info.get("sector") or "") == "Financial Services",
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
    with open(os.path.join(OUT_DIR, "meta.json"), "w", encoding="utf-8") as f:   # 動作確認用
        json.dump({"updated": out["updated"], "tickers": len(codes), "stocks": len(stocks),
                   "with_roe": sum(1 for s in stocks if s["roe"] is not None),
                   "with_roa": sum(1 for s in stocks if s["roa"] is not None),
                   "last_daily": dates["d"][-1], "size_mb": round(os.path.getsize(path) / 1e6, 2)}, f, ensure_ascii=False)


if __name__ == "__main__":
    main()
