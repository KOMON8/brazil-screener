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


def resample_ohlc(df, rule):
    """日足を週足・月足のローソク足に変換"""
    agg = {"Open": "first", "High": "max", "Low": "min", "Close": "last"}
    try:
        out = df[list(agg)].resample(rule).agg(agg)
    except ValueError:
        out = df[list(agg)].resample(rule.replace("ME", "M")).agg(agg)
    return out.dropna()


def write_ohlc(prices, dirpath, dec=2):
    """チャート用のローソク足を銘柄ごとのファイルに書き出す（タップした銘柄だけ読み込む）"""
    os.makedirs(dirpath, exist_ok=True)
    rnd = lambda x: None if pd.isna(x) else round(float(x), dec)
    for c, df in prices.items():
        rec = {}
        for tf, rule in (("d", None), ("w", "W-FRI"), ("m", "ME")):
            o = (df[["Open", "High", "Low", "Close"]] if rule is None else resample_ohlc(df, rule)).tail(BARS[tf])
            rec[tf] = {"t": [d.strftime("%Y-%m-%d") for d in o.index],
                       "o": [rnd(x) for x in o["Open"]], "h": [rnd(x) for x in o["High"]],
                       "l": [rnd(x) for x in o["Low"]], "c": [rnd(x) for x in o["Close"]]}
        with open(os.path.join(dirpath, c + ".json"), "w", encoding="utf-8") as f:
            json.dump(rec, f, separators=(",", ":"))
    print(f"ローソク足: {dirpath}（{len(prices)} 銘柄）")


def fetch_fundamentals(codes, suffix=".SA", cache_path=None, want_roa=True):
    """ROE などの財務の数字（ヤフーの銘柄情報）。前回の記録が新しい銘柄は取り直さない"""
    import time
    import yfinance as yf
    today = dt.date.today()
    cache_path = cache_path or FUND_CACHE
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, encoding="utf-8") as f:
                cache = json.load(f)
        except Exception:
            cache = {}
    todo = [c for c in codes
            if c not in cache or (want_roa and "roa" not in cache[c]) or (today - dt.date.fromisoformat(cache[c].get("checked", "2000-01-01"))).days >= FUND_DAYS]
    print(f"財務の数字を取得中（{len(todo)} 銘柄、残りは記録を使う）...")
    def reset_session(strategy=None):
        """ヤフーの認証（crumb / cookie）を取り直す"""
        try:
            from yfinance.data import YfData
            d = YfData()
            if strategy and hasattr(d, "_set_cookie_strategy"):
                try:
                    d._set_cookie_strategy(strategy)
                except Exception:
                    pass
            for k in ("_crumb", "_cookie"):
                if hasattr(d, k):
                    setattr(d, k, None)
            if hasattr(d, "_set_session"):
                import requests
                try:
                    from curl_cffi import requests as creq
                    d._set_session(creq.Session(impersonate="chrome"))
                except Exception:
                    d._set_session(requests.Session())
        except Exception as e:
            print("  認証のやり直しに失敗:", e)

    reset_session()
    fails = 0
    resets = 0
    for i, c in enumerate(todo):
        if fails >= 10:                    # 続けて失敗したら、認証を取り直して少し待つ
            if resets >= 4:
                print("  取得の失敗が続いたので、残りは次回に取得します")
                break
            resets += 1
            print(f"  失敗が続いたので認証を取り直します（{resets}回目）")
            time.sleep(30 * resets)
            reset_session("csrf" if resets % 2 else "basic")
            fails = 0
        rec = {"checked": today.isoformat()}
        try:
            info = yf.Ticker(c + suffix).info
            roe = info.get("returnOnEquity")
            rec["roe"] = r2(roe * 100) if isinstance(roe, (int, float)) else None
            roa = info.get("returnOnAssets")
            rec["roa"] = r2(roa * 100) if isinstance(roa, (int, float)) else None
            rec["sector"] = info.get("sector") or ""
            fails = 0 if (roe is not None or info.get("shortName")) else fails + 1
        except Exception as e:
            fails += 1
            rec["err"] = str(e)[:80]
            if fails == 1:
                print("  エラー:", c, str(e)[:120])
            rec["checked"] = (today - dt.timedelta(days=FUND_DAYS)).isoformat()   # 失敗したら次回また取る
        cache[c] = rec
        time.sleep(0.4)
        if (i + 1) % 100 == 0:
            print(f"  {i + 1} / {len(todo)}")
    try:
        with open(cache_path, "w", encoding="utf-8") as f:
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
    write_ohlc(prices, os.path.join(OUT_DIR, "ohlc", "br"), 2)
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
