# Apps Script のバックアップ手順

このディレクトリには「SKIHEIL OB合宿管理（年次テンプレート）」に紐づく
Google Apps Script のソースコードを置く。**現時点では未収録。**

## なぜ自動で取得できないのか

スプレッドシートに紐づくスクリプトは**コンテナバインドスクリプト**で、Drive API の
ファイル一覧には現れず、内容を読み出すこともできない。
スプレッドシートを削除するとスクリプトも道連れで消えるため、
**ここにコードを置いておかないとバックアップとして不完全**になる。

## 手順 A: 手でコピーする（確実・おすすめ）

1. スプレッドシートを開き、**拡張機能 → Apps Script** を開く
2. 左側のファイル一覧にある `.gs` / `.html` を 1 つずつ開き、全選択してコピーする
3. このディレクトリに同名のファイルとして保存する（例: `コード.gs`, `index.html`）
4. マニフェストも残す
   - エディタ左の **⚙ プロジェクトの設定** →「`appsscript.json` マニフェスト ファイルを
     エディタで表示する」にチェック
   - 表示された `appsscript.json` をコピーして保存する
5. コミットする前に、**コード内にパスワード・API キー・スプレッドシート ID・
   デプロイ URL が直書きされていないか確認する**（このリポジトリは public）

## 手順 B: clasp を使う

```sh
npm install -g @google/clasp
clasp login

# スクリプト ID は Apps Script エディタの「⚙ プロジェクトの設定 → スクリプト ID」で確認
clasp clone <スクリプトID> --rootDir ./docs/ob-camp-template/apps-script
```

`clasp clone` で作られる `.clasp.json` にはスクリプト ID が入るので、
**public リポジトリにはコミットしない**こと（`.gitignore` 済み）。

## コミット前のチェック

```sh
grep -rn -E "AKfycb|script\.google\.com/macros|PASSWORD|password|[0-9A-Za-z_-]{40,}" \
  docs/ob-camp-template/apps-script/
```

ヒットした場合は、値を `<REDACTED>` に置き換えるか、リポジトリを private にする。
