/**
 * TemplateOps.gs  —  SKIHEIL OB合宿管理シート「テンプレート運用」モジュール
 *
 * 目的:
 *   マスターテンプレートをコピーして合宿ごとの管理シートを作る運用を支援する。
 *   コピー直後に1回だけ実行すれば、前の合宿のデータが消え、
 *   「使い方」シートのWebアプリURLがこのコピーのものに差し替わる。
 *
 * 導入手順:
 *   1. スクリプトエディタで「ファイルを追加 > スクリプト」を選び、名前を TemplateOps にして
 *      このファイルの中身を貼り付ける。
 *   2. 既存の onOpen(...) 関数の末尾に次の1行を追加する。
 *          tplOnOpen();
 *      （onOpen を2つ定義してはいけないため、必ず既存関数に足すこと）
 *      ※ 既存コードに触れたくない場合は、代わりに tplInstallOpenTrigger を1回実行する。
 *         ただしトリガーはコピーに引き継がれないため、コピーごとに再実行が必要になる。
 *
 * 既存コードとの識別子衝突を避けるため、名前はすべて tpl / TPL_ で始めている。
 *
 * WebアプリURLについて:
 *   ScriptApp.getService().getUrl() は、正式なデプロイが無くても Apps Script が内部的に持つ
 *   HEADデプロイのURLを返す。そのIDは「新しいデプロイ」で発行されるIDとは別物で、
 *   /exec を付けても公開URLとしては機能しない（「ファイルを開くことができません」になる）。
 *   そのため自動検出値はあくまで候補として提示し、実際に書き込む値は
 *   「デプロイ > デプロイを管理」に表示されているURLを人が貼り付ける方式にしている。
 */

const TPL_MENU_NAME          = 'テンプレート運用';
const TPL_SHEET_USAGE        = '使い方';
const TPL_SHEET_SETTINGS     = '年度設定';
const TPL_SHEET_EXCEPTIONS   = '料金例外';
const TPL_SHEET_PARTICIPANTS = '参加者';
const TPL_SHEET_DAILY        = '日別集計';

/** 既存の onOpen から呼び出す（またはインストール型トリガーの対象にする）。 */
function tplOnOpen() {
  SpreadsheetApp.getUi()
    .createMenu(TPL_MENU_NAME)
    .addItem('この合宿として初期化（年度設定の内容で）', 'tplShowInitDialog')
    .addItem('WebアプリURLを使い方シートへ反映', 'tplSyncWebAppUrl')
    .addSeparator()
    .addItem('参加者データだけ消去', 'tplClearParticipantsOnly')
    .addToUi();
}

/** onOpen を編集したくない場合の代替。そのコピーに対して1回だけ実行する。 */
function tplInstallOpenTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'tplOnOpen'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tplOnOpen').forSpreadsheet(ss).onOpen().create();
  SpreadsheetApp.getUi().alert('メニュー「' + TPL_MENU_NAME + '」を登録しました。シートを再読み込みしてください。');
}

/* ------------------------------------------------------------------ *
 * 初期化フロー
 * ------------------------------------------------------------------ */

function tplShowInitDialog() {
  const ui  = SpreadsheetApp.getUi();
  const cur = tplCurrentSettings_();
  const participants  = tplCountDataRows_(TPL_SHEET_PARTICIPANTS, '登録ID');
  const exceptionRows = tplCountDataRows_(TPL_SHEET_EXCEPTIONS, '日付');

  const confirm = ui.alert('この内容で初期化します',
    '「年度設定」に入っている内容で初期化します。\n\n'
    + '合宿名: ' + cur.CAMP_NAME + '\n'
    + '実施年度: ' + cur.EVENT_YEAR + '\n'
    + '日程: ' + cur.START_DATE + ' 〜 ' + cur.END_DATE + '\n\n'
    + '違う場合はキャンセルし、「年度設定」の黄色セルを直してから実行し直してください。\n\n'
    + '消去されるもの:\n'
    + '・参加者 ' + participants + ' 行\n'
    + '・日別集計 すべて\n'
    + '・料金例外 ' + exceptionRows + ' 行\n\n'
    + '「リフト設定」は上の日程に合わせて作り直します。\n'
    + 'WebアプリURLは変更しません（メニューの「WebアプリURLを使い方シートへ反映」で登録します）。\n\n'
    + 'この操作は元に戻せません。',
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  let report;
  try {
    report = tplRunInit({
      campName: cur.CAMP_NAME,
      year: String(cur.EVENT_YEAR),
      startDate: cur.START_DATE,
      endDate: cur.END_DATE,
      clearExceptions: true,
      webAppUrl: ''
    });
  } catch (err) {
    ui.alert('初期化できませんでした', String(err && err.message ? err.message : err), ui.ButtonSet.OK);
    return;
  }
  ui.alert('初期化が完了しました', report, ui.ButtonSet.OK);
}

/**
 * デプロイ画面に表示されているウェブアプリURLを入力させる。
 * 空欄・キャンセル・形式不正で諦めた場合は空文字を返す（URLは変更しない）。
 */
function tplAskWebAppUrl_(ui, title) {
  const candidate = tplDetectedWebAppUrl_();
  const current   = tplCurrentUsageUrl_();
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = ui.prompt(title || 'WebアプリURLの登録',
      'Apps Script の「デプロイ > デプロイを管理」に表示されている、\n'
      + '/exec で終わるURLを貼り付けてください。\n\n'
      + (current   ? '現在の記載: ' + current + '\n' : '')
      + (candidate ? '自動検出の候補: ' + candidate
                   + '\n※ 候補が正しいとは限りません。必ずデプロイ画面の値と見比べてください。\n' : '')
      + '\n空欄のまま OK を押すと、URLは変更しません。',
      ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return '';
    const val = String(res.getResponseText() || '').trim();
    if (val === '') return '';
    if (tplIsWebAppUrl_(val)) return val;
    ui.alert('URLの形式が違います',
      'https://script.google.com/macros/s/～/exec の形式で入力してください。\n\n入力された値:\n' + val,
      ui.ButtonSet.OK);
  }
  return '';
}

/* ------------------------------------------------------------------ *
 * 初期化本体
 * ------------------------------------------------------------------ */

function tplRunInit(payload) {
  if (!payload.campName) throw new Error('合宿名を入力してください。');
  if (!/^\d{4}$/.test(payload.year)) throw new Error('実施年度は4桁の数字で入力してください。');
  const start = tplParseYmd_(payload.startDate, '開始日');
  const end   = tplParseYmd_(payload.endDate, '終了日');
  if (start.getTime() > end.getTime()) throw new Error('開始日が終了日より後になっています。');

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('シートが他の処理でロックされています。少し待って再実行してください。');

  try {
    tplSetSetting_('CAMP_NAME',  payload.campName);
    tplSetSetting_('EVENT_YEAR', payload.year);
    tplSetSetting_('START_DATE', payload.startDate);
    tplSetSetting_('END_DATE',   payload.endDate);

    const lines = [];
    lines.push('年度設定を更新しました。');
    lines.push('参加者: ' + tplClearBelowHeader_(TPL_SHEET_PARTICIPANTS, '登録ID') + ' 行を消去');
    lines.push('日別集計: ' + tplClearBelowHeader_(TPL_SHEET_DAILY, '日付') + ' 行を消去');
    if (payload.clearExceptions) {
      lines.push('料金例外: ' + tplClearBelowHeader_(TPL_SHEET_EXCEPTIONS, '日付') + ' 行を消去');
    }

    const url = String(payload.webAppUrl || '').trim();
    if (url) {
      const n = tplWriteUsageUrl_(url);
      lines.push(n > 0
        ? 'WebアプリURLを更新しました:\n' + url
        : '「使い方」シートに書き込み先が見つかりませんでした。手動で貼り付けてください:\n' + url);
    } else {
      lines.push('WebアプリURLは変更していません。\n'
        + 'デプロイ後にメニュー「WebアプリURLを使い方シートへ反映」で登録してください。');
    }

    let liftMsg;
    try {
      syncLiftPlanWithConfig_(getConfig_());
      liftMsg = 'リフト設定を新しい日程で作り直しました。';
    } catch (err) {
      liftMsg = '⚠ リフト設定を同期できませんでした: ' + (err && err.message ? err.message : err)
              + '\n年度設定を確認のうえ、メニュー「SKIHEIL 管理」→「リフト日程を同期」を実行してください。';
    }
    lines.push(liftMsg);

    SpreadsheetApp.flush();
    lines.push('');
    lines.push('次の操作:「リフト設定」で日ごとのスキー場（片品／尾瀬岩鞍／丸沼／購入なし）を指定し、「年度設定」の料金を確認してください。');
    return lines.join('\n');
  } finally {
    lock.releaseLock();
  }
}

function tplClearParticipantsOnly() {
  const ui = SpreadsheetApp.getUi();
  const n = tplCountDataRows_(TPL_SHEET_PARTICIPANTS, '登録ID');
  const res = ui.alert('参加者データの消去',
    '「参加者」シートの ' + n + ' 行と「日別集計」を消去します。元に戻せません。よろしいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  tplClearBelowHeader_(TPL_SHEET_PARTICIPANTS, '登録ID');
  tplClearBelowHeader_(TPL_SHEET_DAILY, '日付');
  ui.alert('参加者と日別集計を消去しました。');
}

function tplSyncWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const url = tplAskWebAppUrl_(ui);
  if (!url) {
    ui.alert('WebアプリURLは変更していません。');
    return;
  }
  const n = tplWriteUsageUrl_(url);
  ui.alert(n > 0
    ? '「使い方」シートのWebアプリURLを更新しました。\n\n' + url
    : '「使い方」シートに書き込み先が見つかりませんでした。以下を手動で貼り付けてください。\n\n' + url);
}

/* ------------------------------------------------------------------ *
 * 内部ユーティリティ
 * ------------------------------------------------------------------ */

function tplSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」が見つかりません。');
  return sh;
}

/** 年度設定シートの「キー」列と「値」列の位置を特定する。 */
function tplSettingsLayout_() {
  const sh = tplSheet_(TPL_SHEET_SETTINGS);
  const values = sh.getDataRange().getValues();
  for (let r = 0; r < Math.min(values.length, 20); r++) {
    let keyCol = -1, valCol = -1;
    for (let c = 0; c < values[r].length; c++) {
      const cell = String(values[r][c] == null ? '' : values[r][c]).trim();
      if (keyCol < 0 && cell.indexOf('キー') === 0) keyCol = c;
      if (valCol < 0 && cell === '値') valCol = c;
    }
    if (keyCol >= 0 && valCol >= 0) {
      return { sheet: sh, values: values, headerRow: r, keyCol: keyCol, valCol: valCol };
    }
  }
  throw new Error('「年度設定」シートで「キー」「値」の見出し行を特定できませんでした。');
}

function tplSettingCell_(key) {
  const L = tplSettingsLayout_();
  for (let r = L.headerRow + 1; r < L.values.length; r++) {
    if (String(L.values[r][L.keyCol]).trim() === key) {
      return L.sheet.getRange(r + 1, L.valCol + 1);
    }
  }
  throw new Error('「年度設定」シートにキー ' + key + ' の行がありません。');
}

function tplCurrentSettings_() {
  const out = {};
  ['CAMP_NAME', 'EVENT_YEAR', 'START_DATE', 'END_DATE'].forEach(function (k) {
    let v = '';
    try { v = tplSettingCell_(k).getValue(); } catch (err) { v = ''; }
    out[k] = (v instanceof Date) ? tplFormatYmd_(v) : String(v == null ? '' : v);
  });
  return out;
}

/** 既存セルの型（日付／数値／文字列）を保ったまま値を書き込む。 */
function tplSetSetting_(key, value) {
  const range = tplSettingCell_(key);
  const cur = range.getValue();
  if (cur instanceof Date) {
    range.setValue(tplParseYmd_(value, key));
  } else if (typeof cur === 'number' && /^-?\d+(\.\d+)?$/.test(String(value))) {
    range.setValue(Number(value));
  } else {
    range.setValue(value);
  }
}

function tplParseYmd_(s, label) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s).trim());
  if (!m) throw new Error((label || '日付') + ' は YYYY-MM-DD 形式で入力してください: ' + s);
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) {
    throw new Error((label || '日付') + ' が存在しない日付です: ' + s);
  }
  return d;
}

function tplFormatYmd_(d) {
  return Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

/** シート内で headerKeyword と完全一致するセルを含む見出し行（0始まり）を返す。 */
function tplFindHeaderRow_(sheet, headerKeyword) {
  const last = Math.min(sheet.getLastRow(), 20);
  if (last < 1) return -1;
  const values = sheet.getRange(1, 1, last, Math.max(sheet.getLastColumn(), 1)).getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === headerKeyword) return r;
    }
  }
  return -1;
}

function tplCountDataRows_(sheetName, headerKeyword) {
  let sh;
  try { sh = tplSheet_(sheetName); } catch (err) { return 0; }
  const h = tplFindHeaderRow_(sh, headerKeyword);
  if (h < 0) return 0;
  return Math.max(0, sh.getLastRow() - (h + 1));
}

/** 見出し行より下のデータ行の値だけを消去する（書式・入力規則は残す）。 */
function tplClearBelowHeader_(sheetName, headerKeyword) {
  const sh = tplSheet_(sheetName);
  const h = tplFindHeaderRow_(sh, headerKeyword);
  if (h < 0) throw new Error('シート「' + sheetName + '」で見出し「' + headerKeyword + '」が見つかりません。');
  const n = sh.getLastRow() - (h + 1);
  if (n <= 0) return 0;
  sh.getRange(h + 2, 1, n, Math.max(sh.getLastColumn(), 1)).clearContent();
  return n;
}

/** 公開ウェブアプリURLとして妥当な形かどうか。 */
function tplIsWebAppUrl_(u) {
  return /^https:\/\/script\.google\.com\/(macros|a\/macros\/[^\/]+)\/s\/[A-Za-z0-9_\-]+\/exec$/
    .test(String(u == null ? '' : u).trim());
}

/**
 * 自動検出したウェブアプリURLの「候補」。
 * HEADデプロイのURLが返ることがあり、それは公開URLとしては使えないため、
 * /exec で終わるものだけを候補として返す。候補が正しい保証はない。
 */
function tplDetectedWebAppUrl_() {
  try {
    const u = String(ScriptApp.getService().getUrl() || '').trim();
    return tplIsWebAppUrl_(u) ? u : '';
  } catch (err) {
    return '';
  }
}

/** 「使い方」シートに現在書かれているウェブアプリURL。無ければ空文字。 */
function tplCurrentUsageUrl_() {
  try {
    const sh = tplSheet_(TPL_SHEET_USAGE);
    const last = sh.getLastRow();
    if (last < 1) return '';
    const values = sh.getRange(1, 1, last, Math.max(sh.getLastColumn(), 1)).getValues();
    const re = /https:\/\/script\.google\.com\/\S*?\/(?:exec|dev)/;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const v = values[r][c];
        if (typeof v === 'string' && v) {
          const m = re.exec(v);
          if (m) return m[0];
        }
      }
    }
    return '';
  } catch (err) {
    return '';
  }
}

/**
 * 「使い方」シート内の script.google.com のURLを url に差し替える。
 * 記載が無い場合は「WebアプリURL」という文言の直下セルに書き込む。
 * 戻り値は書き込んだセル数。
 */
function tplWriteUsageUrl_(url) {
  const sh = tplSheet_(TPL_SHEET_USAGE);
  const last = sh.getLastRow();
  if (last < 1) return 0;
  const cols = Math.max(sh.getLastColumn(), 1);
  const values = sh.getRange(1, 1, last, cols).getValues();
  const re = /https:\/\/script\.google\.com\/\S*?\/(?:exec|dev)/g;

  let count = 0;
  let labelRow = -1, labelCol = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < cols; c++) {
      const v = values[r][c];
      if (typeof v !== 'string' || !v) continue;
      re.lastIndex = 0;
      if (re.test(v)) {
        re.lastIndex = 0;
        sh.getRange(r + 1, c + 1).setValue(v.replace(re, url));
        count++;
      } else if (labelRow < 0 && v.trim() === 'WebアプリURL') {
        labelRow = r;
        labelCol = c;
      }
    }
  }
  if (count === 0 && labelRow >= 0 && labelRow + 1 < last) {
    sh.getRange(labelRow + 2, labelCol + 1).setValue(url);
    count = 1;
  }
  return count;
}
