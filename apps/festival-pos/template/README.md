# スプレッドシートの雛形

レジアプリを動かすスプレッドシートの土台。

| ファイル | 内容 |
| --- | --- |
| `文化祭レジ_ベビーカステラ_テンプレート.xlsx` | 6シート（使い方 / ダッシュボード / 商品 / 設定 / 売上 / レジ締め）入りの雛形 |
| `build_template.py` | 上の .xlsx を生成するスクリプト |

---

## 使い方は 2 通り

### A. スクリプトに作らせる（おすすめ）

雛形を使わず、**空のスプレッドシートから始める**方法。

1. 新しい Google スプレッドシートを作る
2. 拡張機能 → Apps Script に [`../src/`](../src/) の 5 ファイルを貼り付ける
3. シートを開き直し、メニュー「**レジ管理 → シートを初期化 / 修復**」を実行

これだけで 5 シートが揃う。`setupSheets()` が作るので、

- Google ネイティブの表示形式・プルダウンが正しく入る
- ダッシュボードに **QUERY を使った商品別・担当者別・時間帯別**の集計が入る
- パスコードが **4 桁でランダム生成**される

という利点がある。下の B より結果が良い。

### B. .xlsx を取り込む

シート構成を先に見たい場合や、オフラインで商品リストを準備したい場合。

1. `文化祭レジ_ベビーカステラ_テンプレート.xlsx` をダウンロードする
2. Google ドライブにアップロードし、Google スプレッドシートとして開く
3. 以降は A の 2〜3 と同じ（初期化を実行すると書式と集計が最終形になる）

> .xlsx には Excel でも開ける数式しか入っていない。
> QUERY を使う商品別・担当者別・時間帯別の集計は、初期化を実行したときに追加される。
> また `設定` シートのパスコードが **`0000`** になっているので、**必ず変更すること**。

---

## 雛形を作り直す

`Code.gs` の見出し（`SALES_HEADERS` など）を変えたら、雛形も作り直す。

```sh
pip install openpyxl
python3 apps/festival-pos/template/build_template.py
```

見出しがコードと一致しているかは次で確認できる。

```sh
python3 - <<'PY'
import openpyxl, re, io
wb = openpyxl.load_workbook('apps/festival-pos/template/文化祭レジ_ベビーカステラ_テンプレート.xlsx')
code = io.open('apps/festival-pos/src/Code.gs', encoding='utf-8').read()
for sheet, const in [('売上','SALES_HEADERS'), ('商品','ITEMS_HEADERS'), ('レジ締め','CLOSING_HEADERS')]:
    block = re.search(r"const %s = \[(.*?)\];" % const, code, re.S).group(1)
    expected = re.findall(r"'([^']+)'", block)
    actual = [wb[sheet].cell(row=4, column=c).value for c in range(1, len(expected)+1)]
    print(sheet, 'OK' if expected == actual else 'MISMATCH')
PY
```
