// Создаёт пустые планы продаж WB и Ozon: шапки, параметры, формулы — без данных.
// Запуск: node scripts/make-templates.js
const path = require('path');
const fs = require('fs');
const { writeXlsx } = require('./lib/xlsx');

const FIRST = 3;
const LAST = 52;

function planSheet(platform, note, periodLabel) {
  const rows = [];
  rows.push([{ v: `План продаж: ${platform}`, s: 6 }]);
  rows.push([{ v: note, s: 7 }]);
  rows.push(['Артикул', 'Название', 'Остатки, шт', 'Оборачиваемость, дней', `Продано ${periodLabel}, шт`, 'Цена до СПП, ₽/шт',
    'Прогноз ноябрь, шт', 'Прогноз декабрь ×рост, шт', 'Товара хватит (1/0)', 'План декабрь, шт',
    'Сумма до СПП, ₽', 'План, шт/день', '% выкупа', 'План в заказах, шт/день'].map((v) => ({ v, s: 1 })));

  for (let r = FIRST + 1; r <= LAST + 1; r++) {
    rows.push([
      { v: '', s: 0 },
      { v: '', s: 0 },
      { v: '', s: 2 },
      { f: `IF(AND(C${r}>0,ISNUMBER(L${r}),L${r}>0),ROUND(C${r}/L${r},1),"")`, s: 2 },
      { v: '', s: 2 },
      { v: '', s: 2 },
      { f: `IF(E${r}="","",ROUND(E${r}/'Параметры'!$B$5*'Параметры'!$B$6,0))`, s: 2 },
      { f: `IF(G${r}="","",ROUND(G${r}*'Параметры'!$B$8,0))`, s: 2 },
      { v: '', s: 2 },
      { f: `IF(H${r}="","",IF(I${r}=1,H${r},MIN(H${r},MAX(0,C${r}-(G${r}-E${r})))))`, s: 4 },
      { f: `IF(J${r}="","",ROUND(J${r}*F${r},0))`, s: 2 },
      { f: `IF(J${r}="","",ROUND(J${r}/'Параметры'!$B$7,2))`, s: 2 },
      { v: '', s: 3 },
      { f: `IF(OR(L${r}="",M${r}=0),"",ROUND(L${r}/M${r},2))`, s: 2 },
    ]);
  }

  return {
    name: `План ${platform}`,
    rows,
    opts: { widths: [14, 34, 11, 15, 14, 12, 12, 14, 10, 14, 16, 11, 9, 14], freeze: 3, freezeCols: 2 },
  };
}

function paramsSheet(platform) {
  const p = [];
  p.push([{ v: 'Параметры расчёта', s: 6 }]);
  p.push([{ v: 'Меняй значения в столбце B — план пересчитается.', s: 7 }]);
  p.push([]);
  p.push([{ v: 'Показатель', s: 1 }, { v: 'Значение', s: 1 }]);
  p.push([{ v: 'Дней в периоде факта', s: 0 }, { v: 15, s: 2 }]);
  p.push([{ v: 'Дней в первом месяце', s: 0 }, { v: 30, s: 2 }]);
  p.push([{ v: 'Дней в плановом месяце', s: 0 }, { v: 31, s: 2 }]);
  p.push([{ v: 'Рост планового месяца к предыдущему', s: 0 }, { v: 2, s: 4 }]);
  p.push([]);
  p.push([{ v: 'Как заполнять', s: 5 }]);
  p.push([{ v: `Платформа: ${platform}. Заполни A–F и M: артикул, название, остаток, продажи за период факта, цену до СПП, % выкупа (десятичной дробью).`, s: 0 }]);
  p.push([{ v: 'Столбец «Товара хватит»: 1 — товар придёт, план не режется остатком; 0 — продаём только остаток.', s: 0 }]);
  p.push([{ v: 'Цифры в шапке столбцов — формулы, их не вводить.', s: 0 }]);
  return { name: 'Параметры', rows: p, opts: { widths: [42, 14] } };
}

function methodSheet() {
  const m = [];
  m.push([{ v: 'Как считается план', s: 6 }]);
  m.push([]);
  const lines = [
    ['Прогноз первого месяца', 'продажи за период ÷ дней периода × дней месяца'],
    ['Прогноз планового месяца', 'прогноз первого месяца × коэффициент роста'],
    ['План с остатками', 'если флаг = 1 — прогноз целиком; иначе min(прогноз; остаток − (прогноз − продано))'],
    ['Сумма', 'план × цена до СПП'],
    ['План в день', 'план ÷ дней в плановом месяце'],
    ['План в заказах', 'план в день ÷ % выкупа'],
    ['Оборачиваемость', 'остаток ÷ план в день'],
  ];
  for (const [a, b] of lines) m.push([{ v: a, s: 5 }, { v: b, s: 0 }]);
  m.push([]);
  m.push([{ v: 'Правила подсорта', s: 5 }]);
  m.push([{ v: 'Категории, которые дошиваются, ставь с флагом 1. Остальные — 0: продаём только остаток. Список категорий — в rules/03-formulas.md.', s: 0 }]);
  return { name: 'Методика', rows: m, opts: { widths: [34, 100] } };
}

(async () => {
  const dir = path.join(__dirname, '..', 'templates');
  fs.mkdirSync(dir, { recursive: true });

  const wb = path.join(dir, 'plan-wb.xlsx');
  await writeXlsx(wb, [
    planSheet('WB', 'Заполни A–F и M из выгрузки WB. Остальные столбцы — формулы.', 'за период'),
    paramsSheet('WB'),
    methodSheet(),
  ]);

  const oz = path.join(dir, 'plan-ozon.xlsx');
  await writeXlsx(oz, [
    planSheet('Ozon', 'Заполни A–F и M из отчёта Ozon. Остальные столбцы — формулы.', 'за период'),
    paramsSheet('Ozon'),
    methodSheet(),
  ]);

  console.log('готово:');
  console.log(' ', wb);
  console.log(' ', oz);
})();
