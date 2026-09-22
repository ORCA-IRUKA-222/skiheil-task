// GAS グローバルのスタブ
global.console = console;
global.SpreadsheetApp = { getUi: () => ({ createMenu: () => ({ addItem(){return this;}, addSeparator(){return this;}, addToUi(){} }) }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock(){} }) };

const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
// const 宣言をトップレベル eval で拾えるよう、モジュール風に包む
const mod = new Function(src + '\n;return {_toSalesRow_, _prices_, _num_, apiSyncTransactions, SALES_HEADERS};')();

let pass = 0, fail = 0;
function eq(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log((ok ? '  PASS ' : '  FAIL ') + label + (ok ? '' : `\n        期待: ${JSON.stringify(expected)}\n        実際: ${JSON.stringify(actual)}`));
  ok ? pass++ : fail++;
}

const P = { set3: 300, set7: 600, single: 120 };

// 端末側と同じ計算（「セット以外は常にバラ単価」）
const clientTotal = (c, p) => c.set3 * p.set3 + c.set7 * p.set7 + c.single * p.single;
const clientPieces = (c) => c.set3 * 3 + c.set7 * 7 + c.single;

console.log('\n── 料金計算（3個¥300 / 7個¥600 / バラ¥120）──');
eq('3個セット×1', clientTotal({set3:1,set7:0,single:0}, P), 300);
eq('7個セット×1', clientTotal({set3:0,set7:1,single:0}, P), 600);
eq('バラ10個は 10×¥120', clientTotal({set3:0,set7:0,single:10}, P), 1200);
eq('7個セット + 3個セット', clientTotal({set3:1,set7:1,single:0}, P), 900);
eq('7個セット + バラ3個', clientTotal({set3:0,set7:1,single:3}, P), 960);
eq('合計個数 7+3+2', clientPieces({set3:1,set7:1,single:2}), 12);

console.log('\n── 6個問題（セット以外は常にバラ単価の帰結）──');
const six = clientTotal({set3:0,set7:0,single:6}, P);
console.log(`  バラ6個 = ¥${six} / 7個セット = ¥${P.set7} → ` +
  (six > P.set7 ? '6個の方が高い。checkSettings が警告する' : '問題なし'));

console.log('\n── サーバー側の再計算（スナップショット単価から組み直す）──');
const sale = {
  id: 'tx-1', deviceId: 'dev', staff: '田中', clientTime: '2026-11-03T10:15:00.000Z',
  kind: 'sale', voidOf: null,
  items: { set3: 1, set7: 2, single: 4 },
  unitPrices: P, priceVersion: 3,
  total: 300 + 1200 + 480, received: 2000, change: 20, note: ''
};
const row = mod._toSalesRow_(sale, new Date('2026-11-03T10:15:02.000Z'));
const H = mod.SALES_HEADERS;
const at = (name) => row[H.indexOf(name)];
eq('種別', at('種別'), '売上');
eq('3個セット', at('3個セット'), 1);
eq('7個セット', at('7個セット'), 2);
eq('バラ(個)', at('バラ(個)'), 4);
eq('合計個数 = 3+14+4', at('合計個数'), 21);
eq('合計金額', at('合計金額'), 1980);
eq('単価スナップショット(バラ)', at('単価(バラ)'), 120);
eq('価格Ver', at('価格Ver'), 3);
eq('検証', at('検証'), 'OK');

console.log('\n── 取消は数量・金額がマイナス ──');
const voidRow = mod._toSalesRow_(Object.assign({}, sale, {
  id: 'tx-2', kind: 'void', voidOf: 'tx-1', received: null, change: null
}), new Date());
const atv = (name) => voidRow[H.indexOf(name)];
eq('種別', atv('種別'), '取消');
eq('取消元取引ID', atv('取消元取引ID'), 'tx-1');
eq('合計金額', atv('合計金額'), -1980);
eq('合計個数', atv('合計個数'), -21);
eq('お預かりは空', atv('お預かり'), '');
console.log(`  売上 ${at('合計金額')} + 取消 ${atv('合計金額')} = ${at('合計金額') + atv('合計金額')} （SUM で純売上になる）`);

console.log('\n── 端末の計算と食い違う場合は「検証」に出る ──');
const bad = mod._toSalesRow_(Object.assign({}, sale, { id: 'tx-3', total: 9999 }), new Date());
eq('検証が不一致を報告', bad[H.indexOf('検証')].indexOf('不一致') === 0, true);

console.log('\n── 価格変更後に届いた「古い価格の会計」は古い単価のまま記録される ──');
const late = mod._toSalesRow_(Object.assign({}, sale, {
  id: 'tx-4', unitPrices: { set3: 250, set7: 500, single: 100 }, priceVersion: 1,
  total: 250 + 1000 + 400
}), new Date());
eq('旧単価で記録', late[H.indexOf('単価(バラ)')], 100);
eq('旧価格の金額', late[H.indexOf('合計金額')], 1650);
eq('旧バージョン', late[H.indexOf('価格Ver')], 1);
console.log('  → 値上げ前に売った分が値上げ後の金額に化けない');

console.log(`\n${fail === 0 ? '✅' : '❌'} ${pass} passed, ${fail} failed\n`);
process.exit(fail === 0 ? 0 : 1);
