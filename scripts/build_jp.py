# -*- coding: utf-8 -*-
"""
日本株（東証プライム・スタンダード・グロースの内国株式）のデータを作る：site/data/jp.json
  1. JPX の上場銘柄一覧から対象の銘柄を取り出す
  2. 直近3か月の株価で 20日平均売買代金を出し、MIN_VAL_OKU 億円以上の銘柄に絞る
  3. 絞った銘柄の 10 年分の株価と ROE（週1回）を取得
"""
import datetime as dt
import io
import json
import os
import sys
import time

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import screener_core as core
from build_data import r2, fetch_fundamentals, write_ohlc, div_fields, BARS

OUT_DIR = os.path.join(HERE, "..", "site", "data")
LIST_URLS = ("https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx",
             "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls")
LIST_CACHE = os.path.join(HERE, "jp_list.json")
FUND_CACHE_JP = os.path.join(HERE, "fund_cache_jp.json")
MIN_VAL_OKU = float(os.environ.get("MIN_VAL_OKU", "1"))     # 対象にする 20日平均売買代金の下限（億円）
MARKETS = ("プライム（内国株式）", "スタンダード（内国株式）", "グロース（内国株式）")
FIN = ("銀行業", "保険業", "証券、商品先物取引業", "その他金融業")


def load_list():
    """JPX の上場銘柄一覧（取れなければ前回の一覧）"""
    import urllib.request
    for url in LIST_URLS:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as r:
                df = pd.read_excel(io.BytesIO(r.read()), dtype=str)
            df = df[df["市場・商品区分"].isin(MARKETS)]
            out = {str(row["コード"]).strip(): {"name": row["銘柄名"], "sector": row.get("33業種区分", "")}
                   for _, row in df.iterrows()}
            print(f"JPX の一覧: {url} → {len(out)} 銘柄")
            if len(out) > 1000:
                with open(LIST_CACHE, "w", encoding="utf-8") as f:
                    json.dump(out, f, ensure_ascii=False)
                return out
        except Exception as e:
            print("JPX の一覧を取得できませんでした:", url, e)
    if os.path.exists(LIST_CACHE):
        with open(LIST_CACHE, encoding="utf-8") as f:
            print("前回の一覧を使います")
            return json.load(f)
    raise SystemExit("銘柄一覧がありません")


def get_prices(codes, period):
    raw = core.download_chunks([c + ".T" for c in codes], period, "1d")
    out = {}
    for c in codes:
        try:
            df = core.clean_prices(raw[c + ".T"][["Open", "High", "Low", "Close", "Volume"]])
            if len(df) >= 25:
                out[c] = df
        except Exception:
            pass
    return out


def main():
    lst = load_list()
    codes = sorted(lst)
    print(f"東証の内国株式: {len(codes)} 銘柄")

    short = get_prices(codes, "3mo")
    val = {c: (df["Close"] * df["Volume"]).tail(20).mean() / 1e8 for c, df in short.items()}
    target = [c for c in codes if val.get(c, 0) >= MIN_VAL_OKU]
    print(f"20日平均売買代金 {MIN_VAL_OKU:g}億円以上: {len(target)} 銘柄")
    if len(target) < int(os.environ.get("MIN_OK_JP", "100")):
        raise SystemExit("株価をほとんど取得できませんでした（ヤフーに断られた可能性）")

    core.HISTORY_PERIOD = "10y"
    prices = get_prices(target, "10y")
    print(f"10年分の株価を取得できた銘柄: {len(prices)}")
    fund = fetch_fundamentals(list(prices), suffix=".T", cache_path=FUND_CACHE_JP, want_roa=True)

    dates, stocks = {}, []
    series = {c: {"d": df["Close"], "w": core.resample_close(df, "W-FRI"), "m": core.resample_close(df, "ME")}
              for c, df in prices.items()}
    aligned = {c: {} for c in series}
    for tf, n in BARS.items():
        idx = sorted(set().union(*[set(s[tf].index) for s in series.values()]))[-n:]
        dates[tf] = [d.strftime("%Y-%m-%d") for d in idx]
        for c, s in series.items():
            v = s[tf].reindex(idx).ffill(limit=5)
            aligned[c][tf] = [None if pd.isna(x) else round(float(x), 1) for x in v.values]

    for c, df in prices.items():
        cl, v = df["Close"], df["Volume"]
        last, prev = cl.iloc[-1], (cl.iloc[-2] if len(cl) > 1 else None)
        tr = pd.concat([df["High"] - df["Low"], (df["High"] - cl.shift()).abs(), (df["Low"] - cl.shift()).abs()], axis=1).max(axis=1)
        info = lst.get(c, {})
        stocks.append({
            "c": c, "n": info.get("name", ""), "s": info.get("sector", ""),
            "p": r2(last), "chg": r2((last / prev - 1) * 100) if prev else None,
            "val": r2((cl * v).tail(20).mean() / 1e8),          # 20日平均売買代金（億円）
            "atr": r2(tr.tail(14).mean() / last * 100) if last else None,
            "rsi": r2(core.rsi(cl).iloc[-1]),
            "roe": (fund.get(c) or {}).get("roe"),
            "roa": (fund.get(c) or {}).get("roa"),
            **div_fields(fund.get(c), last),
            "fin": info.get("sector", "") in FIN,
            "last": df.index[-1].strftime("%Y-%m-%d"),
            "d": aligned[c]["d"], "w": aligned[c]["w"], "m": aligned[c]["m"],
        })
    stocks.sort(key=lambda s: -(s["val"] or 0))

    os.makedirs(OUT_DIR, exist_ok=True)
    now = dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))
    out = {"updated": now.strftime("%Y-%m-%d %H:%M") + "（日本時間）", "dates": dates, "stocks": stocks}
    write_ohlc(prices, os.path.join(OUT_DIR, "ohlc", "jp"), 1)
    path = os.path.join(OUT_DIR, "jp.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    with open(os.path.join(OUT_DIR, "meta_jp.json"), "w", encoding="utf-8") as f:
        json.dump({"updated": out["updated"], "listed": len(codes), "target": len(target), "stocks": len(stocks),
                   "with_roe": sum(1 for s in stocks if s["roe"] is not None),
                   "with_roa": sum(1 for s in stocks if s["roa"] is not None),
                   "with_div": sum(1 for s in stocks if s["dy"] is not None), "last_daily": dates["d"][-1],
                   "size_mb": round(os.path.getsize(path) / 1e6, 2)}, f, ensure_ascii=False)
    print(f"書き出し: {path}（{os.path.getsize(path) / 1e6:.1f} MB、{len(stocks)} 銘柄）")


if __name__ == "__main__":
    main()
