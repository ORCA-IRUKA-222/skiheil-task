/**
 * 文化祭レジ（ベビーカステラ）— シートの初期化とメンテナンス
 *
 * メニュー「レジ管理 → シートを初期化 / 修復」から実行する。
 * 既存の売上データと商品は消さない（足りないシートと見出しだけを作る）。
 */

const THEME = {
  title: '#1f2937',
  headerBg: '#374151',
  headerFg: '#ffffff',
  editable: '#fff8c4',
  locked: '#f3f4f6',
  band: '#e5e7eb',
  note: '#6b7280'
};

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  _setupSettingsSheet_(ss);
  _setupItemsSheet_(ss);
  _setupSalesSheet_(ss);
  _setupClosingSheet_(ss);
  _setupDashboardSheet_(ss);

  // 表示順を整える
  [SHEETS.DASHBOARD, SHEETS.ITEMS, SHEETS.SETTINGS, SHEETS.SALES, SHEETS.CLOSING]
    .forEach(function (name, i) {
      const sheet = ss.getSheetByName(name);
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    });
  ss.setActiveSheet(ss.getSheetByName(SHEETS.ITEMS));

  const passcode = _readSettings_().PASSCODE;
  SpreadsheetApp.getUi().alert(
    'シートを用意しました。\n\n' +
    'パスコード: ' + passcode + '（「設定」シートで変更できます）\n\n' +
    '1.「商品」シートに味ごとの行を作り、3個・7個・バラの価格を入れる\n' +
    '2.「レジ管理 → 設定をチェック」で価格の妥当性を確認\n' +
    '3. 拡張機能 → Apps Script →「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」\n' +
    '   （次のユーザーとして実行: 自分／アクセスできるユーザー: 全員）\n' +
    '4. 発行された URL をレジ端末で開く'
  );
}

function _sheetOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function _initialSettingValue_(key, defaultValue) {
  // パスコードは推測されないよう毎回ランダムに作る
  if (key === 'PASSCODE') return String(Math.floor(1000 + Math.random() * 9000));
  if (key === 'CATALOG_UPDATED') return new Date();
  return defaultValue;
}

function _titleBlock_(sheet, title, description, width) {
  sheet.getRange(1, 1, 1, width).merge()
    .setValue(title)
    .setFontSize(14).setFontWeight('bold')
    .setFontColor('#ffffff').setBackground(THEME.title)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
  sheet.getRange(2, 1, 1, width).merge()
    .setValue(description)
    .setFontSize(10).setFontColor(THEME.note).setWrap(true);
}

function _headerRow_(sheet, row, headers) {
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontColor(THEME.headerFg)
    .setBackground(THEME.headerBg)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(row);
}

function _bandLabel_(sheet, row, width, text) {
  sheet.getRange(row, 1, 1, width).merge()
    .setValue(text).setFontWeight('bold').setBackground(THEME.band);
}

// ---------- 設定 ----------

function _setupSettingsSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.SETTINGS);
  const width = 4;

  _titleBlock_(sheet, '設定',
    '黄色セルだけを変更します。価格は「商品」シートで管理します。', width);
  _headerRow_(sheet, SETTINGS_HEADER_ROW, ['キー（変更しない）', '設定項目', '値', '説明']);

  const existing = _settingsRowIndex_(sheet);
  SETTINGS_DEFS.forEach(function (def, i) {
    const row = SETTINGS_FIRST_ROW + i;
    const key = def[0];
    sheet.getRange(row, 1).setValue(key);
    sheet.getRange(row, 2).setValue(def[1]);
    // 既にある値は上書きしない（再実行しても設定が飛ばないように）
    if (!existing[key]) {
      sheet.getRange(row, 3).setValue(_initialSettingValue_(key, def[2]));
    }
    sheet.getRange(row, 4).setValue(def[3]);
    sheet.getRange(row, 3).setBackground(def[4] ? THEME.editable : THEME.locked);
  });

  const rows = _settingsRowIndex_(sheet);
  if (rows['CATALOG_UPDATED']) {
    sheet.getRange(rows['CATALOG_UPDATED'], 3).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  if (rows['PASSCODE']) {
    sheet.getRange(rows['PASSCODE'], 3).setNumberFormat('@'); // 先頭0が消えないように文字列扱い
  }

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 340);
}

// ---------- 商品 ----------

function _setupItemsSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.ITEMS);
  const width = ITEMS_HEADERS.length;

  _titleBlock_(sheet, '商品',
    '味ごとに1行つくります。ここを直すと各端末のボタンと価格が自動で入れ替わります（最大15秒）。' +
    '過去の売上金額は会計時の単価で固定されているので変わりません。' +
    '「販売状態」を停止中にすると、その味のボタンがレジから消えます。', width);
  // 注記は必ず見出しより上に置く。データ行の下に書くと、商品として読まれてしまう
  sheet.getRange(3, 1, 1, width).merge()
    .setValue('※「セット以外は常にバラ単価」で計算します。3個・7個ちょうどのときだけセット価格、' +
              'それ以外は 個数 × バラ単価 です。セットとバラは組み合わせられます。')
    .setFontSize(10).setFontColor(THEME.note).setWrap(true);
  sheet.setRowHeight(3, 30);

  _headerRow_(sheet, ITEMS_HEADER_ROW, ITEMS_HEADERS);

  // 以前のバージョンはデータ行の下に注記を書いていた。残っていると ¥0 の商品として
  // 読まれてしまうので消す
  const staleNoteRow = ITEMS_LAST_ROW + 2;
  if (sheet.getMaxRows() >= staleNoteRow + 1) {
    sheet.getRange(staleNoteRow, 1, 2, sheet.getMaxColumns()).breakApart().clearContent();
  }

  // 商品が 1 つも無いときだけサンプルを入れる
  if (_readCatalogSafe_().length === 0) {
    sheet.getRange(ITEMS_FIRST_ROW, 1, ITEMS_SAMPLE.length, width).setValues(ITEMS_SAMPLE);
  }

  const rowCount = ITEMS_LAST_ROW - ITEMS_FIRST_ROW + 1;
  sheet.getRange(ITEMS_FIRST_ROW, 1, rowCount, 1).setBackground(THEME.editable);
  sheet.getRange(ITEMS_FIRST_ROW, 2, rowCount, 3)
    .setBackground(THEME.editable)
    .setNumberFormat('¥#,##0');
  sheet.getRange(ITEMS_FIRST_ROW, 5, rowCount, 1)
    .setBackground(THEME.editable)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(SALE_STATES, true)
        .setAllowInvalid(false)
        .build()
    );

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidths(2, 3, 110);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 260);

}

function _readCatalogSafe_() {
  try {
    return _readCatalog_();
  } catch (err) {
    return [];
  }
}

// ---------- 売上 ----------

function _setupSalesSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.SALES);

  _titleBlock_(sheet, '売上',
    'レジ端末から自動で追記されます。1行 = 1明細（商品×区分）で、同じ取引IDの行がひとつの会計です。' +
    '取引合計・お預かり・お釣り・検証は、その取引の先頭明細行にだけ入ります。' +
    '行の編集・削除はしないでください（端末側と食い違います）。' +
    '打ち間違いはアプリの「取消」で訂正すると、数量と金額がマイナスの取消行として記録されます。',
    SALES_HEADERS.length);
  _headerRow_(sheet, SALES_HEADER_ROW, SALES_HEADERS);

  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('F:F').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ['M', 'N', 'P', 'Q', 'R'].forEach(function (col) {
    sheet.getRange(col + ':' + col).setNumberFormat('¥#,##0');
  });

  sheet.setColumnWidth(1, 145);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 145);
  sheet.setColumnWidth(8, 230);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 90);
}

// ---------- レジ締め ----------

function _setupClosingSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.CLOSING);

  _titleBlock_(sheet, 'レジ締め',
    'アプリの「レジ締め」から記録されます。' +
    '理論現金 = 釣銭準備金 + そのシフトの現金売上。差額がマイナスなら現金が足りません。',
    CLOSING_HEADERS.length);
  _headerRow_(sheet, CLOSING_HEADER_ROW, CLOSING_HEADERS);

  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ['E', 'F', 'G', 'H', 'I'].forEach(function (col) {
    sheet.getRange(col + ':' + col).setNumberFormat('¥#,##0');
  });
  sheet.setColumnWidth(1, 145);
  sheet.setColumnWidth(4, 145);
}

// ---------- ダッシュボード ----------

function _setupDashboardSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.DASHBOARD);
  const width = 8;

  _titleBlock_(sheet, 'ダッシュボード',
    '「売上」シートから自動集計されます。取消は数量・金額がマイナスで入るため、そのまま合計すれば純売上になります。',
    width);

  const cards = [
    ['売上金額', "=IFERROR(SUM('売上'!$N$5:$N),0)", '¥#,##0'],
    ['販売個数', "=IFERROR(SUM('売上'!$L$5:$L),0)", '#,##0"個"'],
    ['会計件数', "=IFERROR(COUNTIFS('売上'!$C$5:$C,1,'売上'!$G$5:$G,\"売上\"),0)", '#,##0"件"'],
    ['取消件数', "=IFERROR(COUNTIFS('売上'!$C$5:$C,1,'売上'!$G$5:$G,\"取消\"),0)", '#,##0"件"']
  ];
  cards.forEach(function (card, i) {
    const col = i * 2 + 1;
    sheet.getRange(4, col, 1, 2).merge()
      .setValue(card[0]).setFontWeight('bold')
      .setBackground(THEME.headerBg).setFontColor(THEME.headerFg)
      .setHorizontalAlignment('center');
    sheet.getRange(5, col, 2, 2).merge()
      .setFormula(card[1]).setNumberFormat(card[2])
      .setFontSize(18).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // 商品別（味ごと）
  _bandLabel_(sheet, 8, width, '商品別（味ごと）');
  sheet.getRange(9, 1).setFormula(
    "=IFERROR(QUERY('売上'!$A$5:$T," +
    "\"select I, sum(K), sum(L), sum(N) where I is not null group by I " +
    "label I '商品名', sum(K) '数量', sum(L) '販売個数', sum(N) '売上金額'\",0)," +
    "\"データがありません\")"
  );

  // 区分別（3個セット / 7個セット / バラ）
  _bandLabel_(sheet, 20, width, '区分別');
  sheet.getRange(21, 1, 1, 3).setValues([['区分', '数量', '売上金額']])
    .setFontWeight('bold').setBackground('#f3f4f6');
  ['3個セット', '7個セット', 'バラ'].forEach(function (label, i) {
    const row = 22 + i;
    sheet.getRange(row, 1).setValue(label);
    sheet.getRange(row, 2).setFormula(
      "=IFERROR(SUMIFS('売上'!$K$5:$K,'売上'!$J$5:$J,\"" + label + "\"),0)"
    ).setNumberFormat('#,##0');
    sheet.getRange(row, 3).setFormula(
      "=IFERROR(SUMIFS('売上'!$N$5:$N,'売上'!$J$5:$J,\"" + label + "\"),0)"
    ).setNumberFormat('¥#,##0');
  });

  // 担当者別
  _bandLabel_(sheet, 26, width, '担当者別');
  sheet.getRange(27, 1).setFormula(
    "=IFERROR(QUERY('売上'!$A$5:$T," +
    "\"select E, sum(L), sum(N) where E is not null group by E " +
    "label E '担当者', sum(L) '販売個数', sum(N) '売上金額'\",0)," +
    "\"データがありません\")"
  );

  // 時間帯別
  _bandLabel_(sheet, 38, width, '時間帯別');
  sheet.getRange(39, 1).setFormula(
    "=IFERROR(QUERY('売上'!$A$5:$T," +
    "\"select hour(F), sum(L), sum(N) where F is not null group by hour(F) " +
    "label hour(F) '時', sum(L) '販売個数', sum(N) '売上金額'\",0)," +
    "\"データがありません\")"
  );

  sheet.setColumnWidths(1, width, 110);
  sheet.setRowHeight(5, 26);
  sheet.setRowHeight(6, 26);
}

// ---------- メンテナンス ----------

function checkSettings() {
  const config = _readSettings_();
  const catalog = _readCatalog_();
  const problems = [];

  if (!String(config.PASSCODE || '').trim()) {
    problems.push('・パスコードが空です。誰でもレジを開けてしまいます');
  }
  if (!String(config.SHOP_NAME || '').trim()) {
    problems.push('・店名が空です');
  }
  if (catalog.length === 0) {
    problems.push('・「商品」シートに商品が 1 つもありません');
  } else if (!catalog.some(function (item) { return item.active; })) {
    problems.push('・販売中の商品が 1 つもありません。すべて停止中になっています');
  }

  // 商品シートの重複名。_readCatalog_ は先勝ちで除いてしまうので、生の行を見る
  const itemsSheet = _sheet_(SHEETS.ITEMS);
  if (itemsSheet.getLastRow() >= ITEMS_FIRST_ROW) {
    const names = itemsSheet
      .getRange(ITEMS_FIRST_ROW, 1, itemsSheet.getLastRow() - ITEMS_FIRST_ROW + 1, 1)
      .getValues()
      .map(function (row) { return String(row[0] || '').trim(); })
      .filter(Boolean);
    const seen = {};
    names.forEach(function (name) {
      if (seen[name] === 1) problems.push('・商品名「' + name + '」が重複しています。下の行は無視されます');
      seen[name] = (seen[name] || 0) + 1;
    });
  }

  catalog.forEach(function (item) {
    const tag = '・' + item.name + ': ';
    if (!(item.set3 > 0) || !(item.set7 > 0) || !(item.single > 0)) {
      problems.push(tag + '価格に 0 以下か空欄があります');
      return;
    }
    if (item.single * 3 <= item.set3) {
      problems.push(tag + '3個セット(¥' + item.set3 + ') が バラ3個(¥' + item.single * 3 + ') 以上です。セットが割高です');
    }
    if (item.single * 7 <= item.set7) {
      problems.push(tag + '7個セット(¥' + item.set7 + ') が バラ7個(¥' + item.single * 7 + ') 以上です。セットが割高です');
    }
    // 「セット以外は常にバラ単価」なので、6個がセット7個より高くなる可能性がある
    if (item.single * 6 > item.set7) {
      problems.push(tag + 'バラ6個(¥' + item.single * 6 + ') が 7個セット(¥' + item.set7 + ') より高くなります。' +
                    '6個のお客さんには7個セットを案内してください');
    }
    if (item.single * 2 > item.set3) {
      problems.push(tag + 'バラ2個(¥' + item.single * 2 + ') が 3個セット(¥' + item.set3 + ') より高くなります');
    }
  });

  const summary = catalog.map(function (item) {
    return '  ' + (item.active ? '●' : '○') + ' ' + item.name +
      '  3個 ¥' + item.set3 + ' / 7個 ¥' + item.set7 + ' / バラ ¥' + item.single;
  }).join('\n');

  SpreadsheetApp.getUi().alert(
    problems.length === 0
      ? '問題は見つかりませんでした。\n\n' + summary +
        '\n\nカタログバージョン ' + config.CATALOG_VERSION
      : '次の点を確認してください。\n\n' + problems.join('\n') + '\n\n---\n' + summary
  );
}

function exportSalesCsv() {
  const sheet = _sheet_(SHEETS.SALES);
  const last = sheet.getLastRow();
  if (last < SALES_FIRST_ROW) {
    SpreadsheetApp.getUi().alert('売上データがまだありません。');
    return;
  }
  const values = sheet
    .getRange(SALES_HEADER_ROW, 1, last - SALES_HEADER_ROW + 1, SALES_HEADERS.length)
    .getDisplayValues();
  const csv = values.map(function (row) {
    return row.map(function (cell) {
      const s = String(cell === null || cell === undefined ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');

  const name = '売上_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm') + '.csv';
  const file = DriveApp.createFile(name, '﻿' + csv, MimeType.CSV);
  SpreadsheetApp.getUi().alert(
    'マイドライブに書き出しました。\n\n' + name + '\n\n' + file.getUrl()
  );
}
