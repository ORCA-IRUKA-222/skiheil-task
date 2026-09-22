// 味ごとの料金計算と、売上明細行の組み立てを検証する
global.SpreadsheetApp = { getUi: () => ({ createMenu: () => ({ addItem(){return this;}, addSeparator(){return this;}, addToUi(){} }) }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock(){} }) };

const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
const mod = new Function(src + '\n;return {_toSalesRows_, _num_, SALES_HEADERS, UNIT_PIECES, UNIT_LABELS};')();

let pass = 0, fail = 0;
function eq(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log((ok ? '  PASS ' : '  FAIL ') + label + (ok ? '' : `\n        期待: ${JSON.stringify(expected)}\n        実際: ${JSON.stringify(actual)}`));
  ok ? pass++ : fail++;
}

// 味ごとに価格が違うカタログ
const CATALOG = {
  'プレーン': { set3: 300, set7: 600, single: 120 },
  'チョコ':   { set3: 350, set7: 700, single: 140 },
  '抹茶':     { set3: 400, set7: 800, single: 160 },
};

// 端末側と同じ計算（「セット以外は常にバラ単価」）
const lineAmount = (l) => l.qty * l.unitPrice;
const total = (lines) => lines.reduce((s, l) => s + lineAmount(l), 0);
const pieces = (lines) => lines.reduce((s, l) => s + l.qty * mod.UNIT_PIECES[l.unit], 0);
const line = (item, unit, qty) => ({ item, unit, qty, unitPrice: CATALOG[item][unit] });

console.log('\n── 味ごとの単価が正しく使われる ──');
eq('プレーン 3個セット', total([line('プレーン','set3',1)]), 300);
eq('チョコ 3個セット',   total([line('チョコ','set3',1)]), 350);
eq('抹茶 7個セット',     total([line('抹茶','set7',1)]), 800);
eq('チョコ バラ10個',    total([line('チョコ','single',10)]), 1400);

console.log('\n── 味をまたいだ会計 ──');
const mixed = [line('プレーン','set7',1), line('チョコ','set3',2), line('抹茶','single',4)];
eq('プレーン7個 + チョコ3個×2 + 抹茶バラ4', total(mixed), 600 + 700 + 640);
eq('合計個数 7 + 6 + 4', pieces(mixed), 17);

console.log('\n── 6個問題は味ごとに起きる（checkSettings が警告する）──');
Object.keys(CATALOG).forEach((name) => {
  const p = CATALOG[name];
  const six = p.single * 6;
  console.log(`  ${name}: バラ6個 ¥${six} / 7個セット ¥${p.set7} → ` +
    (six > p.set7 ? '6個の方が高い' : '問題なし'));
});

console.log('\n── サーバー側の明細行（1行 = 1明細）──');
const tx = {
  id: 'tx-1', deviceId: 'dev', staff: '田中', clientTime: '2026-11-03T10:15:00.000Z',
  kind: 'sale', voidOf: null, lines: mixed, catalogVersion: 3,
  total: total(mixed), received: 2000, change: 2000 - total(mixed), note: ''
};
const built = mod._toSalesRows_(tx, new Date('2026-11-03T10:15:02.000Z'));
const H = mod.SALES_HEADERS;
const col = (name) => H.indexOf(name);
const at = (row, name) => row[col(name)];

eq('明細は3行', built.rows.length, 3);
eq('明細番号', built.rows.map((r) => at(r, '明細番号')), [1, 2, 3]);
eq('商品名', built.rows.map((r) => at(r, '商品名')), ['プレーン', 'チョコ', '抹茶']);
eq('区分', built.rows.map((r) => at(r, '区分')), ['7個セット', '3個セット', 'バラ']);
eq('数量', built.rows.map((r) => at(r, '数量')), [1, 2, 4]);
eq('個数', built.rows.map((r) => at(r, '個数')), [7, 6, 4]);
eq('単価スナップショット', built.rows.map((r) => at(r, '単価')), [600, 350, 160]);
eq('金額', built.rows.map((r) => at(r, '金額')), [600, 700, 640]);
eq('取引合計は先頭行だけ', built.rows.map((r) => at(r, '取引合計')), [1940, '', '']);
eq('お預かりは先頭行だけ', built.rows.map((r) => at(r, 'お預かり')), [2000, '', '']);
eq('お釣りは先頭行だけ', built.rows.map((r) => at(r, 'お釣り')), [60, '', '']);
eq('検証は先頭行だけ', built.rows.map((r) => at(r, '検証')), ['OK', '', '']);
eq('金額の合計 = 取引合計', built.rows.reduce((s, r) => s + at(r, '金額'), 0), 1940);

console.log('\n── 取消は数量・金額がマイナス ──');
const voided = mod._toSalesRows_(Object.assign({}, tx, {
  id: 'tx-2', kind: 'void', voidOf: 'tx-1', received: null, change: null
}), new Date());
eq('種別', voided.rows.map((r) => at(r, '種別')), ['取消', '取消', '取消']);
eq('取消元取引ID', at(voided.rows[0], '取消元取引ID'), 'tx-1');
eq('数量がマイナス', voided.rows.map((r) => at(r, '数量')), [-1, -2, -4]);
eq('金額がマイナス', voided.rows.map((r) => at(r, '金額')), [-600, -700, -640]);
eq('取引合計がマイナス', at(voided.rows[0], '取引合計'), -1940);
eq('取消にお預かりは入らない', at(voided.rows[0], 'お預かり'), '');
console.log('  売上 1940 + 取消 -1940 = 0 （SUM で純売上になる）');

console.log('\n── 端末の計算と食い違う場合は「検証」に出る ──');
const bad = mod._toSalesRows_(Object.assign({}, tx, { id: 'tx-3', total: 9999 }), new Date());
eq('検証が不一致を報告', at(bad.rows[0], '検証').indexOf('不一致') === 0, true);

console.log('\n── 値上げ後に届いた「古い価格の会計」は古い単価のまま ──');
const late = mod._toSalesRows_({
  id: 'tx-4', staff: '佐藤', clientTime: '2026-11-03T09:00:00.000Z', kind: 'sale',
  lines: [{ item: 'チョコ', unit: 'set7', qty: 2, unitPrice: 500 }],
  catalogVersion: 1, total: 1000, received: 1000, change: 0
}, new Date());
eq('旧単価で記録', at(late.rows[0], '単価'), 500);
eq('旧価格の金額', at(late.rows[0], '金額'), 1000);
eq('旧バージョン', at(late.rows[0], 'カタログVer'), 1);
console.log('  → 値上げ前に売った分が値上げ後の金額に化けない');

console.log('\n── 不正な明細は落とす ──');
const dirty = mod._toSalesRows_({
  id: 'tx-5', kind: 'sale', total: 300, catalogVersion: 3,
  lines: [
    { item: 'プレーン', unit: 'set3', qty: 1, unitPrice: 300 },
    { item: '', unit: 'set3', qty: 1, unitPrice: 300 },        // 商品名なし
    { item: 'チョコ', unit: 'set9', qty: 1, unitPrice: 999 },  // 未知の区分
    { item: '抹茶', unit: 'single', qty: 0, unitPrice: 160 }   // 数量0
  ]
}, new Date());
eq('有効な1明細だけ残る', dirty.rows.length, 1);
eq('残ったのはプレーン', at(dirty.rows[0], '商品名'), 'プレーン');

const empty = mod._toSalesRows_({ id: 'tx-6', kind: 'sale', lines: [] }, new Date());
eq('明細ゼロなら行を作らない', empty.rows.length, 0);
eq('理由を返す', empty.error.length > 0, true);

console.log(`\n${fail === 0 ? '✅' : '❌'} ${pass} passed, ${fail} failed\n`);
process.exit(fail === 0 ? 0 : 1);
