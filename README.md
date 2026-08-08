# 日本のひと月の買いもの

経済産業省「商業動態統計調査」の公開データを、そのまま図にして誰でも読めるようにしたサイト。

## 中身

```
data/raw/          e-Stat から落とした公表Excel（拡張子は .xls だが中身は xlsx）
build/build_data.py  Excel → site/data.json への変換
site/index.html    サイト本体（外部ライブラリなし・グラフは自前SVG）
site/data.json     表示用データ（build_data.py の出力）
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

### つまずきどころ

- **www.meti.go.jp は curl の HTTP/2 で 403**（CloudFront の bot 判定）。`--http1.1` で通る。
- **拡張子 `.xls` だが中身は xlsx**。`openpyxl.load_workbook(io.BytesIO(...))` で読む。パス文字列を渡すと拡張子を見て誤判定する。
- **単位がファイルで違う**。大規模卸売店（000031387993）だけ億円、他は百万円。
- 欠測は `***`。

## 更新

```bash
python3 build/build_data.py   # data/raw の Excel から site/data.json を作り直す
```

新しい月が出たら Excel を落とし直してから実行する。

## 出典・ライセンス

- 出典：「商業動態統計調査」（経済産業省）
- 出典：政府統計の総合窓口（e-Stat）https://www.e-stat.go.jp/
- e-Stat のコンテンツは政府標準利用規約（第2.0版）＝ CC BY 互換。出典を書けば商用も含め自由に使える。加工した場合は加工した旨の明記が要る（サイト内に記載済み）。

このサイトは個人が作ったもので、経済産業省・政府が作成・監修したものではない。
