# 日本のひと月の買いもの

経済産業省「商業動態統計調査」の公開データを、そのまま図にして誰でも読めるようにしたサイト。

## 中身

```
data/raw/            商業動態統計の公表Excel（拡張子は .xls だが中身は xlsx）
data/kigyo/          企業活動基本調査の公表Excel（こちらは本物の旧BIFF形式）
data/kigyo/xlsx_converted/  上記を LibreOffice で xlsx にしたもの
build/build_data.py  商業動態統計 → site/data.json
build/build_kigyo.py 企業活動基本調査 → site/kigyo.json
site/index.html      サイト本体（外部ライブラリなし・グラフは自前SVG）
site/data.json       買い物のデータ
site/kigyo.json      企業ベンチマークのデータ
```

## データの取り方

e-Stat のファイルダウンロードURLを叩くだけ。**HTTP/2 だと 403 になるので `--http1.1` が要る**。

```bash
curl -sL --http1.1 -A "Mozilla/5.0" \
  "https://www.e-stat.go.jp/stat-search/file-download?statInfId=000031387994&fileKind=0" \
  -o data/raw/000031387994.xls
```

| statInfId | 中身 |
|---|---|
| 000031387994 | 百貨店・スーパー 商品別販売額（1980年〜・百万円） |
| 000031387995 | コンビニエンスストア 販売額等（1997年〜・百万円） |
| 000031387996 | 家電大型専門店（2014年〜・百万円） |
| 000031387997 | ドラッグストア（2014年〜・百万円） |
| 000031387998 | ホームセンター（2014年〜・百万円） |

### 企業活動基本調査（企業ベンチマーク用）

| statInfId | 中身 |
|---|---|
| 000040464647 | 第1表 産業別 総括表（企業数・従業者数・売上高・営業利益・付加価値額。161業種） |
| 000040464650 | 第4表 産業別×資本金規模別 |
| 000040464656 | 第10表 産業別 研究開発費（売上高研究開発費比率は計算済みで載っている） |

### つまずきどころ

- **www.meti.go.jp は curl の HTTP/2 で 403**（CloudFront の bot 判定）。`--http1.1` で通る。
- **同じ e-Stat でもファイル形式が違う**。商業動態統計は拡張子 `.xls` だが中身は xlsx（`openpyxl.load_workbook(io.BytesIO(...))` で読む）。企業活動基本調査は本物の旧BIFF形式で openpyxl では読めないので、先に変換する:
  `soffice --headless --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir <out> <file>.xls`
- **単位がファイルで違う**。大規模卸売店（000031387993）は億円、業種別商業販売額（000031387992）は10億円、他は百万円。
- 欠測・秘匿は `***` や `X`。
- 企業活動基本調査は**5年度分が縦に積まれている**ので、年度列で絞る必要がある。

## 更新

```bash
python3 build/build_data.py    # 商業動態統計 → site/data.json
python3 build/build_kigyo.py   # 企業活動基本調査 → site/kigyo.json
```

新しい月・年度が出たら Excel を落とし直してから実行する。

## 出典・ライセンス

- 出典：「商業動態統計調査」（経済産業省）
- 出典：「経済産業省企業活動基本調査」（経済産業省）
- 出典：政府統計の総合窓口（e-Stat）https://www.e-stat.go.jp/
- e-Stat のコンテンツは政府標準利用規約（第2.0版）＝ CC BY 互換。出典を書けば商用も含め自由に使える。加工した場合は加工した旨の明記が要る（サイト内に記載済み）。

このサイトは個人が作ったもので、経済産業省・政府が作成・監修したものではない。

## 公開先

https://imai-design.github.io/nihon-no-kaimono/ （GitHub Pages・gh-pages ブランチ）

```bash
git push origin main
git subtree push --prefix site origin gh-pages
```
