/**
 * 文化祭レジ（ベビーカステラ）— シートの初期化とメンテナンス
 *
 * メニュー「レジ管理 → シートを初期化 / 修復」から実行する。
 * 既存の売上データは消さない（足りないシートと見出しだけを作る）。
 */

const THEME = {
  title: '#1f2937',
  headerBg: '#374151',
  headerFg: '#ffffff',
  editable: '#fff8c4',
  locked: '#f3f4f6',
  note: '#6b7280'
};

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  _setupSettingsSheet_(ss);
  _setupSalesSheet_(ss);
  _setupClosingSheet_(ss);
  _setupDashboardSheet_(ss);

  // 表示順を整える
  [SHEETS.DASHBOARD, SHEETS.SETTINGS, SHEETS.SALES, SHEETS.CLOSING].forEach(function (name, i) {
    const sheet = ss.getSheetByName(name);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));

  const passcode = _readSettings_().PASSCODE;
  SpreadsheetApp.getUi().alert(
    'シートを用意しました。\n\n' +
    'パスコード: ' + passcode + '（「設定」シートで変更できます）\n\n' +
    '1.「設定」シートの黄色セルで価格を設定\n' +
    '2. 拡張機能 → Apps Script →「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」\n' +
    '   （次のユーザーとして実行: 自分／アクセスできるユーザー: 全員）\n' +
    '3. 発行された URL をレジ端末で開く'
  );
}

/**
 * 初期化時に入れる値。パスコードは推測されないよう毎回ランダムに作る。
 */
function _initialSettingValue_(key, defaultValue) {
  if (key === 'PASSCODE') {
    return String(Math.floor(1000 + Math.random() * 9000));
  }
  if (key === 'PRICE_UPDATED') {
    return new Date();
  }
  return defaultValue;
}

function _sheetOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
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

// ---------- 設定 ----------

function _setupSettingsSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.SETTINGS);
  const width = 4;

  _titleBlock_(sheet, '設定',
    '黄色セルだけを変更します。価格を変えると自動で「価格バージョン」が上がり、' +
    '各端末が次の同期（最大30秒）で新しい価格を読み込みます。' +
    '過去の売上金額は、そのとき適用した単価で固定されているので変わりません。', width);
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
  PRICE_KEYS.forEach(function (key) {
    if (rows[key]) sheet.getRange(rows[key], 3).setNumberFormat('¥#,##0');
  });
  if (rows['PRICE_UPDATED']) {
    sheet.getRange(rows['PRICE_UPDATED'], 3).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  if (rows['PASSCODE']) {
    sheet.getRange(rows['PASSCODE'], 3).setNumberFormat('@'); // 先頭0が消えないように文字列扱い
  }

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 340);

  const noteRow = SETTINGS_FIRST_ROW + SETTINGS_DEFS.length + 1;
  sheet.getRange(noteRow, 1, 1, width).merge()
    .setValue('※「セット以外は常にバラ単価」で計算します。' +
              '3個・7個ちょうどのときだけセット価格、それ以外は 個数 × バラ単価 です。')
    .setFontSize(10).setFontColor(THEME.note).setWrap(true);
}

// ---------- 売上 ----------

function _setupSalesSheet_(ss) {
  const sheet = _sheetOrCreate_(ss, SHEETS.SALES);

  _titleBlock_(sheet, '売上',
    'レジ端末から自動で追記されます。行の編集・削除はしないでください（端末側と食い違います）。' +
    '打ち間違いはアプリの「取消」で訂正すると、数量と金額がマイナスの取消行として記録されます。',
    SALES_HEADERS.length);
  _headerRow_(sheet, SALES_HEADER_ROW, SALES_HEADERS);

  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('E:E').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ['L', 'M', 'N', 'P', 'Q', 'R'].forEach(function (col) {
    sheet.getRange(col + ':' + col).setNumberFormat('¥#,##0');
  });

  sheet.setColumnWidth(1, 145);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 145);
  sheet.setColumnWidth(7, 230);
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
    ['売上金額', "=IFERROR(SUM('売上'!$P$5:$P),0)", '¥#,##0'],
    ['販売個数', "=IFERROR(SUM('売上'!$K$5:$K),0)", '#,##0"個"'],
    ['会計件数', "=IFERROR(COUNTIF('売上'!$F$5:$F,\"売上\"),0)", '#,##0"件"'],
    ['取消件数', "=IFERROR(COUNTIF('売上'!$F$5:$F,\"取消\"),0)", '#,##0"件"']
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

  const items = [
    ['3個セット', "=IFERROR(SUM('売上'!$H$5:$H),0)"],
    ['7個セット', "=IFERROR(SUM('売上'!$I$5:$I),0)"],
    ['バラ(個)', "=IFERROR(SUM('売上'!$J$5:$J),0)"],
    ['検証NG', "=IFERROR(COUNTIFS('売上'!$S$5:$S,\"<>OK\",'売上'!$S$5:$S,\"<>\"),0)"]
  ];
  items.forEach(function (item, i) {
    const col = i * 2 + 1;
    sheet.getRange(8, col, 1, 2).merge()
      .setValue(item[0]).setFontWeight('bold')
      .setBackground('#e5e7eb').setHorizontalAlignment('center');
    sheet.getRange(9, col, 1, 2).merge()
      .setFormula(item[1]).setNumberFormat('#,##0')
      .setFontSize(14).setHorizontalAlignment('center');
  });

  sheet.getRange(11, 1, 1, width).merge()
    .setValue('担当者別').setFontWeight('bold').setBackground('#e5e7eb');
  sheet.getRange(12, 1).setFormula(
    "=IFERROR(QUERY('売上'!$A$5:$T," +
    "\"select D, count(B), sum(K), sum(P) where D is not null group by D label count(B) '会計件数', sum(K) '販売個数', sum(P) '売上金額', D '担当者'\",0)," +
    "\"データがありません\")"
  );

  sheet.getRange(22, 1, 1, width).merge()
    .setValue('時間帯別').setFontWeight('bold').setBackground('#e5e7eb');
  sheet.getRange(23, 1).setFormula(
    "=IFERROR(QUERY('売上'!$A$5:$T," +
    "\"select hour(E), count(F), sum(K), sum(P) where E is not null group by hour(E) label hour(E) '時', count(F) '会計件数', sum(K) '販売個数', sum(P) '売上金額'\",0)," +
    "\"データがありません\")"
  );

  sheet.setColumnWidths(1, width, 110);
  sheet.setRowHeight(5, 26);
  sheet.setRowHeight(6, 26);
}

// ---------- メンテナンス ----------

function checkSettings() {
  const config = _readSettings_();
  const problems = [];

  PRICE_KEYS.forEach(function (key) {
    const v = Number(config[key]);
    if (!isFinite(v) || v <= 0) problems.push('・' + key + ' が未入力か 0 以下です');
  });
  if (!String(config.PASSCODE || '').trim()) {
    problems.push('・PASSCODE が空です。誰でもレジを開けてしまいます');
  }
  if (!String(config.SHOP_NAME || '').trim()) {
    problems.push('・SHOP_NAME が空です');
  }

  const single = Number(config.PRICE_SINGLE);
  const set3 = Number(config.PRICE_SET3);
  const set7 = Number(config.PRICE_SET7);
  if (isFinite(single) && isFinite(set3) && single * 3 <= set3) {
    problems.push('・3個セット(¥' + set3 + ') が バラ3個(¥' + single * 3 + ') 以上です。セットが割高になっています');
  }
  if (isFinite(single) && isFinite(set7) && single * 7 <= set7) {
    problems.push('・7個セット(¥' + set7 + ') が バラ7個(¥' + single * 7 + ') 以上です。セットが割高になっています');
  }
  // 「セット以外は常にバラ単価」なので、6個がセット7個より高くなる可能性がある
  if (isFinite(single) && isFinite(set7) && single * 6 > set7) {
    problems.push('・バラ6個(¥' + single * 6 + ') が 7個セット(¥' + set7 + ') より高くなります。' +
                  'お客さんに指摘されやすいので価格を見直すか、6個のお客さんには7個セットを案内してください');
  }
  if (isFinite(single) && isFinite(set3) && single * 2 > set3) {
    problems.push('・バラ2個(¥' + single * 2 + ') が 3個セット(¥' + set3 + ') より高くなります');
  }

  SpreadsheetApp.getUi().alert(
    problems.length === 0
      ? '問題は見つかりませんでした。\n\n' +
        '3個セット ¥' + set3 + ' / 7個セット ¥' + set7 + ' / バラ1個 ¥' + single + '\n' +
        '価格バージョン ' + config.PRICE_VERSION
      : '次の点を確認してください。\n\n' + problems.join('\n')
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
