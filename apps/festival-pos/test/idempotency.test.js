// 冪等性とカタログ配信を、偽のスプレッドシートで検証する
const salesRows = [];                    // 売上シートの 5 行目以降
const SETTINGS = [
  ['SHOP_NAME','店名','ベビーカステラ'],
  ['PASSCODE','パスコード','1234'],
  ['CATALOG_VERSION','カタログバージョン',7],
  ['CATALOG_UPDATED','最終更新日時',new Date('2026-11-03T08:00:00Z')],
];
const ITEMS = [
  ['プレーン', 300, 600, 120, '販売中', ''],
  ['チョコ',   350, 700, 140, '販売中', ''],
  ['抹茶',     400, 800, 160, '停止中', '焼き上がり待ち'],
  ['チョコ',   999, 999, 999, '販売中', '重複行。無視される'],
  ['',         0,   0,   0,   '',       ''],
  ['※「セット以外は常にバラ単価」で計算します。', 0, 0, 0, '', ''],  // 旧版がデータ行の下に書いた注記
];

const readOnly = (data) => ({
  getLastRow: () => 4 + data.length,
  getRange: (r, c, numRows, numCols) => ({
    getValues: () => data.slice(r - 5, r - 5 + numRows).map((row) => row.slice(c - 1, c - 1 + numCols)),
  }),
});
const salesSheet = {
  getLastRow: () => (salesRows.length ? 4 + salesRows.length : 4),
  getRange(r, c, numRows, numCols) {
    return {
      getValues: () => salesRows.slice(r - 5, r - 5 + numRows).map((row) => row.slice(c - 1, c - 1 + numCols)),
      setValues(vals) { vals.forEach((v, i) => { salesRows[r - 5 + i] = v; }); },
    };
  },
};

global.SpreadsheetApp = {
  getUi: () => ({ createMenu: () => ({ addItem(){return this;}, addSeparator(){return this;}, addToUi(){} }) }),
  getActiveSpreadsheet: () => ({
    getSheetByName: (n) =>
      n === '売上' ? salesSheet : n === '設定' ? readOnly(SETTINGS) : n === '商品' ? readOnly(ITEMS) : null,
  }),
};
let lockCalls = 0;
global.LockService = { getScriptLock: () => ({ tryLock: () => { lockCalls++; return true; }, releaseLock(){} }) };

const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
const mod = new Function(src + '\n;return {apiSyncTransactions, apiBootstrap, _readCatalog_, SALES_HEADERS};')();

let pass = 0, fail = 0;
function eq(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log((ok ? '  PASS ' : '  FAIL ') + label + (ok ? '' : `  期待:${JSON.stringify(expected)} 実際:${JSON.stringify(actual)}`));
  ok ? pass++ : fail++;
}
const tx = (id, item, qty) => ({
  id, deviceId: 'dev-a', staff: '田中', clientTime: '2026-11-03T10:00:00.000Z',
  kind: 'sale', lines: [{ item, unit: 'set3', qty, unitPrice: item === 'チョコ' ? 350 : 300 }],
  catalogVersion: 7, total: (item === 'チョコ' ? 350 : 300) * qty, received: 1000, change: 0,
});

console.log('\n── カタログ配信 ──');
const boot = mod.apiBootstrap('1234');
eq('店名', boot.shopName, 'ベビーカステラ');
eq('カタログバージョン', boot.catalogVersion, 7);
eq('商品は3件（重複と空行を除く）', boot.catalog.length, 3);
eq('味ごとの価格', boot.catalog.map((i) => [i.name, i.set3, i.set7, i.single]),
  [['プレーン',300,600,120], ['チョコ',350,700,140], ['抹茶',400,800,160]]);
eq('重複行は先勝ち（チョコは350のまま）', boot.catalog[1].set3, 350);
eq('停止中も active:false で返す', boot.catalog.map((i) => i.active), [true, true, false]);
eq('価格のない注記の行は商品にしない', boot.catalog.some((i) => i.name.indexOf('※') === 0), false);
eq('間違ったパスコードは拒否', mod.apiBootstrap('9999').ok, false);

console.log('\n── パスコード ──');
eq('間違ったパスコードでは同期できない', mod.apiSyncTransactions('9999', [tx('a','プレーン',1)]).ok, false);
eq('拒否時は1行も書かれない', salesRows.length, 0);

console.log('\n── 通常の送信 ──');
let r = mod.apiSyncTransactions('1234', [tx('a','プレーン',1), tx('b','チョコ',2)]);
eq('2件受理', r.acceptedIds, ['a','b']);
eq('売上シートに2行', salesRows.length, 2);
eq('カタログも一緒に返る', r.catalog.length, 3);
eq('カタログバージョンも返る', r.catalogVersion, 7);

console.log('\n── オフライン再送（同じ取引IDを送り直す）──');
r = mod.apiSyncTransactions('1234', [tx('a','プレーン',1), tx('b','チョコ',2), tx('c','抹茶',1)]);
eq('3件とも受理として返る', r.acceptedIds, ['a','b','c']);
eq('行は3行のまま（a,b は二重計上されない）', salesRows.length, 3);

console.log('\n── まるごと再送（全件重複）──');
r = mod.apiSyncTransactions('1234', [tx('a','プレーン',1), tx('b','チョコ',2), tx('c','抹茶',1)]);
eq('全件受理として返る（端末のキューが掃ける）', r.acceptedIds, ['a','b','c']);
eq('行は増えない', salesRows.length, 3);

console.log('\n── 複数明細の取引でも取引ID単位で冪等 ──');
const multi = {
  id: 'd', deviceId: 'dev-b', staff: '佐藤', clientTime: '2026-11-03T10:05:00.000Z',
  kind: 'sale', catalogVersion: 7, total: 600 + 700, received: 1500, change: 200,
  lines: [
    { item: 'プレーン', unit: 'set7', qty: 1, unitPrice: 600 },
    { item: 'チョコ',   unit: 'set7', qty: 1, unitPrice: 700 },
  ],
};
mod.apiSyncTransactions('1234', [multi]);
eq('2明細ぶん増える', salesRows.length, 5);
mod.apiSyncTransactions('1234', [multi]);
eq('再送しても増えない', salesRows.length, 5);

console.log('\n── 同一バッチ内に同じIDが混ざった場合 ──');
mod.apiSyncTransactions('1234', [tx('e','プレーン',1), tx('e','プレーン',1)]);
eq('1行だけ追加', salesRows.length, 6);

console.log('\n── 明細のない取引は受理せず理由を返す ──');
r = mod.apiSyncTransactions('1234', [{ id: 'f', kind: 'sale', lines: [], total: 0 }]);
eq('受理しない', r.acceptedIds.indexOf('f'), -1);
eq('rejected に入る', r.rejected.map((x) => x.id), ['f']);
eq('理由がある', r.rejected[0].reason.length > 0, true);
eq('行は増えない', salesRows.length, 6);

console.log('\n── 空バッチ（カタログのポーリング）はロックを取らない ──');
const before = lockCalls;
r = mod.apiSyncTransactions('1234', []);
eq('ロック未取得', lockCalls, before);
eq('カタログは返る', r.catalogVersion, 7);

console.log('\n── 集計 ──');
const H = mod.SALES_HEADERS;
const amount = H.indexOf('金額');
// a:プレーン300 / b:チョコ350×2 / c:抹茶(txヘルパーは300) / multi:600+700 / e:プレーン300
eq('金額の合計', salesRows.reduce((s, row) => s + row[amount], 0), 300 + 700 + 300 + 600 + 700 + 300);

console.log(`\n${fail===0?'✅':'❌'} ${pass} passed, ${fail} failed\n`);
process.exit(fail===0?0:1);
