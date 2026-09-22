// ── 冪等性テスト：偽のスプレッドシートでサーバー API を動かす ──
const rows = [];                       // 売上シートの 5 行目以降
const SETTINGS = [
  ['PRICE_SET3','3個セット',300],['PRICE_SET7','7個セット',600],
  ['PRICE_SINGLE','バラ1個',120],['SHOP_NAME','店名','ベビーカステラ'],
  ['PASSCODE','パスコード','1234'],['PRICE_VERSION','価格バージョン',3],
];

function makeSalesSheet() {
  return {
    getLastRow: () => rows.length ? 4 + rows.length : 4,
    getRange(r, c, numRows, numCols) {
      return {
        getValues: () => rows.slice(r - 5, r - 5 + numRows).map(row => row.slice(c - 1, c - 1 + numCols)),
        setValues(vals) { vals.forEach((v, i) => { rows[r - 5 + i] = v; }); }
      };
    }
  };
}
function makeSettingsSheet() {
  return {
    getLastRow: () => 4 + SETTINGS.length,
    getRange: (r, c, numRows, numCols) => ({
      getValues: () => SETTINGS.slice(r - 5, r - 5 + numRows).map(row => row.slice(c - 1, c - 1 + numCols))
    })
  };
}
const salesSheet = makeSalesSheet();
global.SpreadsheetApp = {
  getUi: () => ({ createMenu: () => ({ addItem(){return this;}, addSeparator(){return this;}, addToUi(){} }) }),
  getActiveSpreadsheet: () => ({
    getSheetByName: (n) => n === '売上' ? salesSheet : n === '設定' ? makeSettingsSheet() : null
  })
};
let lockCalls = 0;
global.LockService = { getScriptLock: () => ({ tryLock: () => { lockCalls++; return true; }, releaseLock(){} }) };

const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
const mod = new Function(src + '\n;return {apiSyncTransactions, SALES_HEADERS};')();

let pass = 0, fail = 0;
function eq(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log((ok ? '  PASS ' : '  FAIL ') + label + (ok ? '' : `  期待:${JSON.stringify(expected)} 実際:${JSON.stringify(actual)}`));
  ok ? pass++ : fail++;
}
const tx = (id, n) => ({
  id, deviceId:'dev-a', staff:'田中', clientTime:'2026-11-03T10:0'+n+':00.000Z',
  kind:'sale', items:{set3:1,set7:0,single:0}, unitPrices:{set3:300,set7:600,single:120},
  priceVersion:3, total:300, received:500, change:200, note:''
});

console.log('\n── パスコード ──');
eq('間違ったパスコードは拒否', mod.apiSyncTransactions('9999', [tx('a',1)]).ok, false);
eq('拒否時は1行も書かれない', rows.length, 0);

console.log('\n── 通常の送信 ──');
let r = mod.apiSyncTransactions('1234', [tx('a',1), tx('b',2)]);
eq('2件受理', r.acceptedIds, ['a','b']);
eq('売上シートに2行', rows.length, 2);
eq('価格も一緒に返る', r.prices, {set3:300,set7:600,single:120});
eq('価格バージョンも返る', r.priceVersion, 3);

console.log('\n── オフライン再送（同じ取引IDを送り直す）──');
r = mod.apiSyncTransactions('1234', [tx('a',1), tx('b',2), tx('c',3)]);
eq('3件とも受理として返る', r.acceptedIds, ['a','b','c']);
eq('行は3行のまま（a,b は二重計上されない）', rows.length, 3);

console.log('\n── まるごと再送（全件重複）──');
r = mod.apiSyncTransactions('1234', [tx('a',1), tx('b',2), tx('c',3)]);
eq('全件受理として返る（端末のキューが掃ける）', r.acceptedIds, ['a','b','c']);
eq('行は増えない', rows.length, 3);

console.log('\n── 同一バッチ内に同じIDが混ざった場合 ──');
r = mod.apiSyncTransactions('1234', [tx('d',4), tx('d',4)]);
eq('1行だけ追加', rows.length, 4);

console.log('\n── 空バッチ（価格ポーリング）はロックを取らない ──');
const before = lockCalls;
r = mod.apiSyncTransactions('1234', []);
eq('ロック未取得', lockCalls, before);
eq('価格は返る', r.priceVersion, 3);

console.log('\n── 合計金額 ──');
const idx = mod.SALES_HEADERS.indexOf('合計金額');
eq('4件 × ¥300', rows.reduce((s, row) => s + row[idx], 0), 1200);

console.log(`\n${fail===0?'✅':'❌'} ${pass} passed, ${fail} failed\n`);
process.exit(fail===0?0:1);
