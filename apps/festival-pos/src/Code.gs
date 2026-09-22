/**
 * 文化祭レジ（ベビーカステラ）— サーバー側
 *
 * ── 同期の設計 ──────────────────────────────────────────────
 *  価格マスタ : スプレッドシート → アプリ の「一方向」
 *  売上明細   : アプリ → スプレッドシート の「一方向・追記のみ」
 *
 *  同じデータが双方向に流れないので、原理的に競合が起きない。
 *  そのうえで次の 2 つを守る。
 *
 *   1. 売上行には「会計時に適用した単価」を値で焼き込む（スナップショット）。
 *      あとから価格を変えても過去の売上金額は動かない。
 *   2. 取引 ID（UUID）で冪等化する。オフライン再送で二重計上しない。
 *
 *  価格変更の検知はタイムスタンプではなく「価格バージョン（単調増加の整数）」で行う。
 *  端末の時計がずれていても壊れないため。
 * ────────────────────────────────────────────────────────────
 */

const SHEETS = {
  SETTINGS: '設定',
  SALES: '売上',
  CLOSING: 'レジ締め',
  DASHBOARD: 'ダッシュボード'
};

/** 設定シートのレイアウト */
const SETTINGS_HEADER_ROW = 4;
const SETTINGS_FIRST_ROW = 5;

/** 設定の既定値（キー, 項目名, 値, 説明, 編集可か） */
const SETTINGS_DEFS = [
  ['PRICE_SET3',    '3個セット',      300,                     'ボタン「3個」の価格', true],
  ['PRICE_SET7',    '7個セット',      600,                     'ボタン「7個」の価格', true],
  ['PRICE_SINGLE',  'バラ1個',        120,                     'セット以外はこの単価 × 個数', true],
  ['SHOP_NAME',     '店名',           'ベビーカステラ',         'アプリ上部に表示', true],
  ['PASSCODE',      'パスコード',      '',                      'アプリを開くときに入力する。初期化時に自動生成', true],
  ['PRICE_VERSION', '価格バージョン',  1,                       '自動更新。手で編集しない', false],
  ['PRICE_UPDATED', '価格更新日時',    '',                      '自動更新。手で編集しない', false]
];

/** 価格そのものを持つキー。ここが編集されたらバージョンを上げる */
const PRICE_KEYS = ['PRICE_SET3', 'PRICE_SET7', 'PRICE_SINGLE'];

/** 売上シートの列見出し */
const SALES_HEADERS = [
  '受信日時', '取引ID', '端末ID', '担当者', '会計日時', '種別', '取消元取引ID',
  '3個セット', '7個セット', 'バラ(個)', '合計個数',
  '単価(3個)', '単価(7個)', '単価(バラ)', '価格Ver',
  '合計金額', 'お預かり', 'お釣り', '検証', '備考'
];
const SALES_HEADER_ROW = 4;
const SALES_FIRST_ROW = 5;
const SALES_COL_TXID = 2;

/** レジ締めシートの列見出し */
const CLOSING_HEADERS = [
  '記録日時', '担当者', '端末ID', '開始時刻',
  '釣銭準備金', '現金売上(理論)', '理論現金', '実査現金', '差額',
  '取引件数', '備考'
];
const CLOSING_HEADER_ROW = 4;

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('レジ管理')
    .addItem('シートを初期化 / 修復', 'setupSheets')
    .addItem('価格バージョンを上げる（全端末に再読込させる）', 'bumpPriceVersionManually')
    .addSeparator()
    .addItem('売上をCSVで書き出す', 'exportSalesCsv')
    .addItem('設定をチェック', 'checkSettings')
    .addToUi();
}

/**
 * 設定シートの価格が編集されたら価格バージョンを上げる。
 * 簡易トリガーなので、スプレッドシートへの書き込みだけを行う。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEETS.SETTINGS) return;
    if (e.range.getColumn() !== 3) return; // 値は C 列

    const rows = _settingsRowIndex_(sheet);
    const touched = PRICE_KEYS.some(function (key) {
      const row = rows[key];
      return row && row >= e.range.getRow() && row <= e.range.getLastRow();
    });
    if (!touched) return;

    _bumpPriceVersion_(sheet, rows);
  } catch (err) {
    // 簡易トリガーで例外を投げるとユーザーに不可解なエラーが出るだけなので握りつぶす
    console.error(err);
  }
}

function bumpPriceVersionManually() {
  const sheet = _sheet_(SHEETS.SETTINGS);
  const version = _bumpPriceVersion_(sheet, _settingsRowIndex_(sheet));
  SpreadsheetApp.getUi().alert(
    '価格バージョンを ' + version + ' にしました。\n' +
    '各端末は次回の同期（最大30秒）で新しい価格を読み込みます。'
  );
}

function _bumpPriceVersion_(sheet, rows) {
  const versionRow = rows['PRICE_VERSION'];
  const updatedRow = rows['PRICE_UPDATED'];
  const current = Number(sheet.getRange(versionRow, 3).getValue()) || 0;
  const next = current + 1;
  sheet.getRange(versionRow, 3).setValue(next);
  if (updatedRow) {
    sheet.getRange(updatedRow, 3).setValue(new Date());
  }
  return next;
}

// ============================================================
// Web アプリ
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('レジ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// アプリから呼ばれる API（google.script.run）
// ============================================================

/**
 * 起動時。パスコードを確認して店名と価格を返す。
 */
function apiBootstrap(passcode) {
  const config = _readSettings_();
  if (String(passcode || '') !== String(config.PASSCODE)) {
    return { ok: false, error: 'パスコードが違います' };
  }
  return {
    ok: true,
    shopName: String(config.SHOP_NAME || 'レジ'),
    prices: _prices_(config),
    priceVersion: Number(config.PRICE_VERSION) || 0,
    serverTime: new Date().toISOString()
  };
}

/**
 * 未送信の取引をまとめて受け取る。
 *
 * 取引 ID で冪等化しているので、同じものを何度送られても二重計上しない。
 * 受理できた ID を返すので、端末側はそれを「送信済み」に変える。
 */
function apiSyncTransactions(passcode, transactions) {
  const config = _readSettings_();
  if (String(passcode || '') !== String(config.PASSCODE)) {
    return { ok: false, error: 'パスコードが違います' };
  }

  const list = Array.isArray(transactions) ? transactions : [];
  const result = {
    ok: true,
    acceptedIds: [],
    prices: _prices_(config),
    priceVersion: Number(config.PRICE_VERSION) || 0,
    serverTime: new Date().toISOString()
  };
  if (list.length === 0) return result;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    return { ok: false, error: '他の端末が書き込み中です。少し待ってから再送してください' };
  }

  try {
    const sheet = _sheet_(SHEETS.SALES);
    const known = _existingTransactionIds_(sheet);
    const now = new Date();
    const rows = [];

    list.forEach(function (tx) {
      if (!tx || !tx.id) return;
      if (known[tx.id]) {
        // すでに記録済み。受理済みとして返すことで端末側の未送信キューを掃除させる
        result.acceptedIds.push(tx.id);
        return;
      }
      known[tx.id] = true;
      rows.push(_toSalesRow_(tx, now));
      result.acceptedIds.push(tx.id);
    });

    if (rows.length > 0) {
      const start = Math.max(sheet.getLastRow() + 1, SALES_FIRST_ROW);
      sheet.getRange(start, 1, rows.length, SALES_HEADERS.length).setValues(rows);
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * レジ締めの記録。
 */
function apiCloseRegister(passcode, payload) {
  const config = _readSettings_();
  if (String(passcode || '') !== String(config.PASSCODE)) {
    return { ok: false, error: 'パスコードが違います' };
  }
  const p = payload || {};
  const float = _num_(p.floatAmount);
  const cashSales = _num_(p.cashSales);
  const expected = float + cashSales;
  const counted = _num_(p.countedCash);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    return { ok: false, error: '他の端末が書き込み中です。少し待ってから再送してください' };
  }
  try {
    const sheet = _sheet_(SHEETS.CLOSING);
    sheet.appendRow([
      new Date(),
      String(p.staff || ''),
      String(p.deviceId || ''),
      p.startedAt ? new Date(p.startedAt) : '',
      float,
      cashSales,
      expected,
      counted,
      counted - expected,
      _num_(p.txCount),
      String(p.note || '')
    ]);
    return { ok: true, expected: expected, difference: counted - expected };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 内部処理
// ============================================================

function _toSalesRow_(tx, receivedAt) {
  const sign = tx.kind === 'void' ? -1 : 1;
  const set3 = Math.round(_num_(tx.items && tx.items.set3));
  const set7 = Math.round(_num_(tx.items && tx.items.set7));
  const single = Math.round(_num_(tx.items && tx.items.single));

  const unit3 = _num_(tx.unitPrices && tx.unitPrices.set3);
  const unit7 = _num_(tx.unitPrices && tx.unitPrices.set7);
  const unitSingle = _num_(tx.unitPrices && tx.unitPrices.single);

  const pieces = set3 * 3 + set7 * 7 + single;
  // 端末が送ってきた単価から金額を組み直す。端末側の計算と食い違えば「検証」列に出す
  const recomputed = set3 * unit3 + set7 * unit7 + single * unitSingle;
  const claimed = _num_(tx.total);
  const verdict = Math.abs(recomputed - claimed) < 0.5
    ? 'OK'
    : '不一致（端末: ' + claimed + ' / 再計算: ' + recomputed + '）';

  return [
    receivedAt,
    String(tx.id),
    String(tx.deviceId || ''),
    String(tx.staff || ''),
    tx.clientTime ? new Date(tx.clientTime) : '',
    tx.kind === 'void' ? '取消' : '売上',
    String(tx.voidOf || ''),
    sign * set3,
    sign * set7,
    sign * single,
    sign * pieces,
    unit3,
    unit7,
    unitSingle,
    _num_(tx.priceVersion),
    sign * recomputed,
    tx.kind === 'void' ? '' : _numOrBlank_(tx.received),
    tx.kind === 'void' ? '' : _numOrBlank_(tx.change),
    verdict,
    String(tx.note || '')
  ];
}

function _existingTransactionIds_(sheet) {
  const last = sheet.getLastRow();
  const map = {};
  if (last < SALES_FIRST_ROW) return map;
  const values = sheet
    .getRange(SALES_FIRST_ROW, SALES_COL_TXID, last - SALES_FIRST_ROW + 1, 1)
    .getValues();
  values.forEach(function (row) {
    const id = String(row[0] || '').trim();
    if (id) map[id] = true;
  });
  return map;
}

function _readSettings_() {
  const sheet = _sheet_(SHEETS.SETTINGS);
  const last = sheet.getLastRow();
  const config = {};
  if (last < SETTINGS_FIRST_ROW) return config;
  const values = sheet
    .getRange(SETTINGS_FIRST_ROW, 1, last - SETTINGS_FIRST_ROW + 1, 3)
    .getValues();
  values.forEach(function (row) {
    const key = String(row[0] || '').trim();
    if (key) config[key] = row[2];
  });
  return config;
}

function _settingsRowIndex_(sheet) {
  const last = sheet.getLastRow();
  const rows = {};
  if (last < SETTINGS_FIRST_ROW) return rows;
  const values = sheet
    .getRange(SETTINGS_FIRST_ROW, 1, last - SETTINGS_FIRST_ROW + 1, 1)
    .getValues();
  values.forEach(function (row, i) {
    const key = String(row[0] || '').trim();
    if (key) rows[key] = SETTINGS_FIRST_ROW + i;
  });
  return rows;
}

function _prices_(config) {
  return {
    set3: _num_(config.PRICE_SET3),
    set7: _num_(config.PRICE_SET7),
    single: _num_(config.PRICE_SINGLE)
  };
}

function _sheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」がありません。メニュー「レジ管理 → シートを初期化 / 修復」を実行してください。');
  }
  return sheet;
}

function _num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function _numOrBlank_(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return isFinite(n) ? n : '';
}
