# brazil-screener

ブラジル株（B3 の株式：普通株・優先株・ユニット）を、日足・週足・月足の **移動平均（MA）の角度** で絞り込み、
選んだ銘柄のチャート（終値と MA）を見るための Web ページです。iPad のブラウザ（Safari）で使えます。

- 毎晩 18:48（日本時間）に GitHub Actions がヤフーファイナンスから株価を取り、ページのデータを更新します
- 角度・絞り込み・チャートはページ側で計算するので、MA の期間などを変えるとその場で結果が変わります
- 条件（投資可能額・売買代金・各足の角度）は、使っている端末に保存されます

## 中身
- `site/index.html` … 画面
- `scripts/build_data.py` … データを作る（`site/data/all.json`）
- `scripts/screener_core.py` … パソコン版スクリーナーと共通の取得・計算部品
- `.github/workflows/update.yml` … 毎晩の自動実行と公開

手動で更新したいときは、Actions → 「毎晩データ更新」→ Run workflow。
