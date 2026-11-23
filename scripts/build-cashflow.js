// Платёжный календарь и ДДС-прогноз из двух чисел: продажи и доли выплат.
// Пример:
// node scripts/build-cashflow.js --oz-sales 16000000 --wb-sales 4000000 --oz-rate 0.52 --wb-rate 0.35 --purchase 3000000 --opening 0 --out output/cashflow.xlsx
const path = require('path');
const fs = require('fs');
const { writeXlsx, colName } = require('./lib/xlsx');

function arg(name, def) {
  const i = process.argv.indexOf(`--${name}`);
  if (i < 0) return def;
  const v = process.argv[i + 1];
  return v === undefined ? def : (Number.isFinite(Number(v)) ? Number(v) : v);
}

const OZ_SALES = arg('oz-sales', 0);
const WB_SALES = arg('wb-sales', 0);
const OZ_RATE = arg('oz-rate', 0.5);
const WB_RATE = arg('wb-rate', 0.35);
const NOV_OZ = arg('nov-oz', 0);
const NOV_WB = arg('nov-wb', 0);
const PURCHASE = arg('purchase', 0);
const OPENING = arg('opening', 0);
const PURCHASE_WEEK = arg('purchase-week', 0);
const OUT = arg('out', path.join(__dirname, '..', 'output', 'cashflow.xlsx'));

// недели: ноябрь (для лагов), декабрь, январь, февраль
const weeks = [
  { label: '01–07.11', days: 7, month: 'nov' },
  { label: '08–14.11', days: 7, month: 'nov' },
  { label: '15–21.11', days: 7, month: 'nov' },
  { label: '22–30.11', days: 9, month: 'nov' },
  { label: '01–07.12', days: 7, month: 'dec' },
  { label: '08–14.12', days: 7, month: 'dec' },
  { label: '15–21.12', days: 7, month: 'dec' },
  { label: '22–31.12', days: 10, month: 'dec' },
  { label: '01–07.01', days: 7, month: 'jan' },
  { label: '08–14.01', days: 7, month: 'jan' },
  { label: '15–21.01', days: 7, month: 'jan' },
  { label: '22–31.01', days: 10, month: 'jan' },
  { label: '01–07.02', days: 7, month: 'feb' },
];

const OZ_LAG = 1;
const WB_LAG = 5;

const ozByWeek = new Array(weeks.length).fill(0);
const wbByWeek = new Array(weeks.length).fill(0);

function split(total, month, daysInMonth) {
  return total ? Math.round((total * daysInMonth) / daysInMonth) : 0;
}

weeks.forEach((w, i) => {
  let sales = 0;
  if (w.month === 'nov') sales = NOV_OZ ? (NOV_OZ * w.days) / 30 : 0;
  if (w.month === 'dec') sales = OZ_SALES ? (OZ_SALES * w.days) / 31 : 0;
  if (sales) ozByWeek[i + OZ_LAG] += Math.round(sales * OZ_RATE);

  let salesWb = 0;
  if (w.month === 'nov') salesWb = NOV_WB ? (NOV_WB * w.days) / 30 : 0;
  if (w.month === 'dec') salesWb = WB_SALES ? (WB_SALES * w.days) / 31 : 0;
  if (salesWb) wbByWeek[i + WB_LAG] = (wbByWeek[i + WB_LAG] || 0) + Math.round(salesWb * WB_RATE);
});

const outWeeks = weeks.map((w, i) => ({ ...w, oz: ozByWeek[i] || 0, wb: wbByWeek[i] || 0 })).slice(4);

(async () => {
  const rows = [];
  rows.push([{ v: 'Платёжный календарь и ДДС-прогноз', s: 6 }]);
  rows.push([{ v: `Ozon: продажи ${OZ_SALES.toLocaleString('ru-RU')} ₽ × ${(OZ_RATE * 100).toFixed(1)}%, лаг ${OZ_LAG} нед. WB: ${WB_SALES.toLocaleString('ru-RU')} ₽ × ${(WB_RATE * 100).toFixed(1)}%, лаг ${WB_LAG} нед. Закупка ${PURCHASE.toLocaleString('ru-RU')} ₽.`, s: 7 }]);
  rows.push([]);

  const head = ['Показатель', ...outWeeks.map((w) => w.label), 'Итого'];
  rows.push(head.map((v) => ({ v, s: 1 })));

  const first = 2;
  const last = first + outWeeks.length - 1;
  const total = last + 1;

  const rowOpen = rows.length + 1;
  rows.push([{ v: 'Остаток на начало (введи фактический)', s: 0 }, { v: OPENING, s: 4 }, ...outWeeks.slice(1).map(() => null), null]);

  const rowOz = rows.length + 1;
  rows.push([{ v: 'Поступления Ozon', s: 0 }, ...outWeeks.map((w) => (w.oz ? { v: w.oz, s: 2 } : null)), { f: `SUM(B${rowOz}:${colName(last)}${rowOz})`, s: 4 }]);

  const rowWb = rows.length + 1;
  rows.push([{ v: 'Поступления WB', s: 0 }, ...outWeeks.map((w) => (w.wb ? { v: w.wb, s: 2 } : null)), { f: `SUM(B${rowWb}:${colName(last)}${rowWb})`, s: 4 }]);

  const rowIn = rows.length + 1;
  rows.push([{ v: 'Итого поступления', s: 5 }, ...outWeeks.map((_, i) => ({ f: `${colName(first + i)}${rowOz}+${colName(first + i)}${rowWb}`, s: 4 })), { f: `SUM(B${rowIn}:${colName(last)}${rowIn})`, s: 4 }]);

  const rowBuy = rows.length + 1;
  const buyCells = outWeeks.map((_, i) => (i === PURCHASE_WEEK ? { v: PURCHASE, s: 2 } : null));
  rows.push([{ v: 'Закупка товара (по себестоимости)', s: 0 }, ...buyCells, { f: `SUM(B${rowBuy}:${colName(last)}${rowBuy})`, s: 4 }]);

  const rowOut = rows.length + 1;
  rows.push([{ v: 'Итого расходы', s: 5 }, ...outWeeks.map((_, i) => ({ f: `${colName(first + i)}${rowBuy}`, s: 4 })), { f: `SUM(B${rowOut}:${colName(last)}${rowOut})`, s: 4 }]);

  const rowNet = rows.length + 1;
  rows.push([{ v: 'Чистый поток', s: 5 }, ...outWeeks.map((_, i) => ({ f: `${colName(first + i)}${rowIn}-${colName(first + i)}${rowOut}`, s: 4 })), { f: `SUM(B${rowNet}:${colName(last)}${rowNet})`, s: 4 }]);

  const rowBal = rows.length + 1;
  rows.push([{ v: 'Остаток на конец', s: 5 }, { f: `B${rowOpen}+B${rowNet}`, s: 4 }, ...outWeeks.slice(1).map((_, i) => ({ f: `${colName(first + i)}${rowBal}+${colName(first + i + 1)}${rowNet}`, s: 4 })), null]);

  rows.push([]);
  rows.push([{ v: 'Как считается', s: 5 }]);
  rows.push([{ v: 'Продажи месяца делятся по дням недели. Поступления = продажи недели × доля выплат, сдвиг на лаг площадки.', s: 0 }]);
  rows.push([{ v: 'Ноябрьские продажи (если переданы) идут с теми же лагами и попадают в декабрь–январь.', s: 0 }]);
  rows.push([{ v: 'Закупка ставится в одну неделю (по умолчанию первая) — поменяй, если оплата разбита.', s: 0 }]);
  rows.push([{ v: 'Прочие расходы (ФОТ, налоги, ФФ, реклама) не включены — добавь строками вручную.', s: 0 }]);

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  await writeXlsx(OUT, [
    { name: 'ДДС', rows, opts: { widths: [34, 13, 13, 13, 13, 13, 13, 13, 13, 13, 16], freeze: 4, freezeCols: 1 } },
  ]);

  const sum = (arr) => arr.reduce((a, b) => a + b, 0);
  console.log('файл:', OUT);
  console.log('Ozon:', sum(outWeeks.map((w) => w.oz)), '| WB:', sum(outWeeks.map((w) => w.wb)), '| закупка:', PURCHASE);
})();
