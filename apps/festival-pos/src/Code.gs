/**
 * 文化祭レジ（ベビーカステラ）— サーバー側
 *
 * ── 同期の設計 ──────────────────────────────────────────────
 *  商品・価格 : スプレッドシート → アプリ の「一方向」
 *  売上明細   : アプリ → スプレッドシート の「一方向・追記のみ」
 *
 *  同じデータが双方向に流れないので、原理的に競合が起きない。
 *  そのうえで次の 2 つを守る。
 *
 *   1. 売上行には「会計時に適用した商品名と単価」を値で焼き込む（スナップショット）。
 *      あとから価格を変えても過去の売上金額は動かない。
 *   2. 取引 ID（UUID）で冪等化する。オフライン再送で二重計上しない。
 *
 *  変更の検知はタイムスタンプではなく「カタログバージョン（単調増加の整数）」で行う。
 *  端末の時計がずれていても壊れないため。
 * ────────────────────────────────────────────────────────────
 */

const SHEETS = {
  SETTINGS: '設定',
  ITEMS: '商品',
  SALES: '売上',
  CLOSING: 'レジ締め',
  DASHBOARD: 'ダッシュボード'
};

/** 設定シートのレイアウト */
const SETTINGS_HEADER_ROW = 4;
const SETTINGS_FIRST_ROW = 5;

/** 設定の既定値（キー, 項目名, 値, 説明, 編集可か） */
const SETTINGS_DEFS = [
  ['SHOP_NAME',        '店名',              'ベビーカステラ', 'アプリ上部に表示', true],
  ['PASSCODE',         'パスコード',         '',              'アプリを開くときに入力する。初期化時に自動生成', true],
  ['CATALOG_VERSION',  'カタログバージョン',  1,               '自動更新。手で編集しない', false],
  ['CATALOG_UPDATED',  '最終更新日時',       '',              '自動更新。手で編集しない', false]
];

/** 自動更新のため、編集されてもバージョンを上げないキー */
const AUTO_KEYS = ['CATALOG_VERSION', 'CATALOG_UPDATED'];

/** 商品シートのレイアウト */
const ITEMS_HEADER_ROW = 4;
const ITEMS_FIRST_ROW = 5;
const ITEMS_LAST_ROW = 54;
const ITEMS_HEADERS = ['商品名', '3個セット', '7個セット', 'バラ1個', '販売状態', '備考'];

/** 初期化時に入れておくサンプル（味ごとに1行） */
const ITEMS_SAMPLE = [
  ['プレーン', 300, 600, 120, '販売中', ''],
  ['チョコ',   350, 700, 140, '販売中', '']
];

const SALE_STATES = ['販売中', '停止中'];

/** 売り方の区分。バラは「セット以外は常にバラ単価」で使う */
const UNITS = ['set3', 'set7', 'single'];
const UNIT_LABELS = { set3: '3個セット', set7: '7個セット', single: 'バラ' };
const UNIT_PIECES = { set3: 3, set7: 7, single: 1 };

/**
 * 売上シートの列見出し。1 行 = 1 明細。
 * 取引単位の値（取引合計・お預かり・お釣り・検証・備考）は、その取引の先頭明細行にだけ入る。
 */
const SALES_HEADERS = [
  '受信日時', '取引ID', '明細番号', '端末ID', '担当者', '会計日時', '種別', '取消元取引ID',
  '商品名', '区分', '数量', '個数', '単価', '金額', 'カタログVer',
  '取引合計', 'お預かり', 'お釣り', '検証', '備考'
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
    .addItem('カタログバージョンを上げる（全端末に再読込させる）', 'bumpCatalogVersionManually')
    .addSeparator()
    .addItem('売上をCSVで書き出す', 'exportSalesCsv')
    .addItem('設定をチェック', 'checkSettings')
    .addToUi();
}

/**
 * 商品や設定が編集されたらカタログバージョンを上げる。
 * 簡易トリガーなので、スプレッドシートへの書き込みだけを行う。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const name = sheet.getName();

    if (name === SHEETS.ITEMS) {
      if (e.range.getLastRow() <= ITEMS_HEADER_ROW) return;
    } else if (name === SHEETS.SETTINGS) {
      if (e.range.getColumn() !== 3) return; // 値は C 列
      const rows = _settingsRowIndex_(sheet);
      const autoOnly = AUTO_KEYS.every(function (key) {
        return rows[key] && rows[key] >= e.range.getRow() && rows[key] <= e.range.getLastRow();
      });
      const touchedAuto = AUTO_KEYS.some(function (key) {
        return rows[key] && rows[key] >= e.range.getRow() && rows[key] <= e.range.getLastRow();
      });
      // 自動更新セルだけが変わった場合は無限に上がり続けないよう何もしない
      if (touchedAuto && autoOnly) return;
    } else {
      return;
    }

    _bumpCatalogVersion_();
  } catch (err) {
    // 簡易トリガーで例外を投げるとユーザーに不可解なエラーが出るだけなので握りつぶす
    console.error(err);
  }
}

function bumpCatalogVersionManually() {
  const version = _bumpCatalogVersion_();
  SpreadsheetApp.getUi().alert(
    'カタログバージョンを ' + version + ' にしました。\n' +
    '各端末は次回の同期（最大15秒）で新しい商品と価格を読み込みます。'
  );
}

function _bumpCatalogVersion_() {
  const sheet = _sheet_(SHEETS.SETTINGS);
  const rows = _settingsRowIndex_(sheet);
  const versionRow = rows['CATALOG_VERSION'];
  const updatedRow = rows['CATALOG_UPDATED'];
  const current = Number(sheet.getRange(versionRow, 3).getValue()) || 0;
  const next = current + 1;
  sheet.getRange(versionRow, 3).setValue(next);
  if (updatedRow) sheet.getRange(updatedRow, 3).setValue(new Date());
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
 * 起動時。パスコードを確認して店名と商品カタログを返す。
 */
function apiBootstrap(passcode) {
  const config = _readSettings_();
  if (!_passcodeOk_(passcode, config)) {
    return { ok: false, error: 'パスコードが違います' };
  }
  return {
    ok: true,
    shopName: String(config.SHOP_NAME || 'レジ'),
    catalog: _readCatalog_(),
    catalogVersion: Number(config.CATALOG_VERSION) || 0,
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
  if (!_passcodeOk_(passcode, config)) {
    return { ok: false, error: 'パスコードが違います' };
  }

  const list = Array.isArray(transactions) ? transactions : [];
  const result = {
    ok: true,
    acceptedIds: [],
    rejected: [],
    catalog: _readCatalog_(),
    catalogVersion: Number(config.CATALOG_VERSION) || 0,
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
      const built = _toSalesRows_(tx, now);
      if (built.rows.length === 0) {
        // 明細のない取引は書かない。端末側で原因がわかるよう理由を返す
        result.rejected.push({ id: String(tx.id), reason: built.error || '明細がありません' });
        return;
      }
      known[tx.id] = true;
      built.rows.forEach(function (row) { rows.push(row); });
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
  if (!_passcodeOk_(passcode, config)) {
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
// 売上行の組み立て
// ============================================================

/**
 * 取引 1 件を明細行の配列にする。1 行 = 1 明細。
 * 取引単位の値は先頭行にだけ入れ、以降の行は空にする。
 */
function _toSalesRows_(tx, receivedAt) {
  const sign = tx.kind === 'void' ? -1 : 1;
  const lines = Array.isArray(tx.lines) ? tx.lines : [];
  const valid = [];

  lines.forEach(function (line) {
    if (!line) return;
    const item = String(line.item || '').trim();
    const unit = String(line.unit || '');
    const qty = Math.round(_num_(line.qty));
    if (!item || UNITS.indexOf(unit) < 0 || qty <= 0) return;
    valid.push({ item: item, unit: unit, qty: qty, unitPrice: _num_(line.unitPrice) });
  });

  if (valid.length === 0) {
    return { rows: [], error: '有効な明細がありません' };
  }

  // 端末が送ってきた単価から金額を組み直す。端末側の計算と食い違えば「検証」列に出す
  const recomputed = valid.reduce(function (sum, l) { return sum + l.qty * l.unitPrice; }, 0);
  const claimed = _num_(tx.total);
  const verdict = Math.abs(recomputed - claimed) < 0.5
    ? 'OK'
    : '不一致（端末: ' + claimed + ' / 再計算: ' + recomputed + '）';

  const rows = valid.map(function (l, index) {
    const first = index === 0;
    return [
      receivedAt,
      String(tx.id),
      index + 1,
      String(tx.deviceId || ''),
      String(tx.staff || ''),
      tx.clientTime ? new Date(tx.clientTime) : '',
      tx.kind === 'void' ? '取消' : '売上',
      String(tx.voidOf || ''),
      l.item,
      UNIT_LABELS[l.unit],
      sign * l.qty,
      sign * l.qty * UNIT_PIECES[l.unit],
      l.unitPrice,
      sign * l.qty * l.unitPrice,
      _num_(tx.priceVersion || tx.catalogVersion),
      first ? sign * recomputed : '',
      first && tx.kind !== 'void' ? _numOrBlank_(tx.received) : '',
      first && tx.kind !== 'void' ? _numOrBlank_(tx.change) : '',
      first ? verdict : '',
      first ? String(tx.note || '') : ''
    ];
  });

  return { rows: rows, error: '' };
}

// ============================================================
// 内部処理
// ============================================================

function _passcodeOk_(given, config) {
  return String(given || '') === String(config.PASSCODE || '');
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

/**
 * 商品シートを読んでカタログにする。
 * 商品名がキー。同じ名前が 2 行あった場合は先に書かれている方を採用する。
 * 停止中の商品も active:false として返す（会計途中のカートの単価が引けなくなるため）。
 */
function _readCatalog_() {
  const sheet = _sheet_(SHEETS.ITEMS);
  const last = sheet.getLastRow();
  if (last < ITEMS_FIRST_ROW) return [];

  const values = sheet
    .getRange(ITEMS_FIRST_ROW, 1, last - ITEMS_FIRST_ROW + 1, ITEMS_HEADERS.length)
    .getValues();

  const seen = {};
  const catalog = [];
  values.forEach(function (row) {
    const name = String(row[0] || '').trim();
    if (!name || seen[name]) return;
    seen[name] = true;
    catalog.push({
      name: name,
      set3: _num_(row[1]),
      set7: _num_(row[2]),
      single: _num_(row[3]),
      active: String(row[4] || '販売中').trim() !== '停止中'
    });
  });
  return catalog;
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
