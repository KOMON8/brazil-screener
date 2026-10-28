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


RB_DROP = -3.0     # 急落日とみなす前日比（%以下）
RB_GAIN = 1.0      # 買値（急落日の終値）からこれ以上（%）高い値を付けたら「戻った」
RB_DAYS = 2        # 急落日の翌日から何日以内に戻ったかを見る
RB_LOOK = 500      # 過去何営業日（約2年）を対象にするか


def rebound_fields(df):
    """急落後の戻り：過去約2年で前日比 RB_DROP% 以下だった日の終値で買ったとして、
    翌日（rb1）・RB_DAYS 日以内（rb）に高値が RB_GAIN% 以上上回った割合、回数（rbn）、平均の最大戻り幅（rbw）"""
    cl, hi = df["Close"].values, df["High"].values
    n = len(cl)
    ev1 = ev2 = 0
    gains = []
    for i in range(max(1, n - RB_LOOK), n - RB_DAYS):
        c0, cp = cl[i], cl[i - 1]
        if not (c0 > 0 and cp > 0) or (c0 / cp - 1) * 100 > RB_DROP:
            continue
        hs = [h for h in hi[i + 1:i + 1 + RB_DAYS] if h == h]
        if len(hs) < RB_DAYS:
            continue
        g = (max(hs) / c0 - 1) * 100
        gains.append(g)
        ev1 += (hs[0] / c0 - 1) * 100 >= RB_GAIN
        ev2 += g >= RB_GAIN
    k = len(gains)
    if not k:
        return {"rbn": 0, "rb": None, "rb1": None, "rbw": None}
    return {"rbn": k, "rb": round(ev2 / k * 100), "rb1": round(ev1 / k * 100), "rbw": r2(float(np.median(gains)))}


def div_fields(f, price):
    """配当：dy＝配当利回り%、dps＝年間1株配当、po＝配当性向%、ex＝権利落ち日"""
    f = f or {}
    d = f.get("div")
    if d is None:
        return {"dy": None, "dps": None, "po": None, "ex": None}
    dy = r2(d / price * 100) if price else None
    if dy is not None and dy > 30:          # 特別配当などで異常に高いものは参考外
        dy = None
    return {"dy": dy, "dps": round(d, 2), "po": f.get("po"), "ex": f.get("ex")}


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
            if c not in cache or (want_roa and "roa" not in cache[c]) or "div" not in cache[c] or (today - dt.date.fromisoformat(cache[c].get("checked", "2000-01-01"))).days >= FUND_DAYS]
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
            div = info.get("dividendRate")
            if not isinstance(div, (int, float)):
                div = info.get("trailingAnnualDividendRate")
            rec["div"] = round(float(div), 4) if isinstance(div, (int, float)) and div > 0 else 0.0   # 年間の1株配当（0＝無配）
            po = info.get("payoutRatio")
            rec["po"] = r2(po * 100) if isinstance(po, (int, float)) else None                      # 配当性向（%）
            ex = info.get("exDividendDate")
            rec["ex"] = dt.datetime.fromtimestamp(ex, dt.timezone.utc).strftime("%Y-%m-%d") if isinstance(ex, (int, float)) and ex > 0 else None
            fails = 0 if (roe is not None or info.get("shortName")) else fails + 1
        except Exception as e:
            fails += 1
            old = {k: v for k, v in (cache.get(c) or {}).items() if k not in ("checked", "err", "div")}
            rec.update(old)                   # 失敗しても前回の ROE などは残す
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
            **div_fields(fund.get(c), last),
            **rebound_fields(df),
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
                   "with_div": sum(1 for s in stocks if s["dy"] is not None),
                   "last_daily": dates["d"][-1], "size_mb": round(os.path.getsize(path) / 1e6, 2)}, f, ensure_ascii=False)


if __name__ == "__main__":
    main()
