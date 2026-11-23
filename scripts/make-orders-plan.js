// Превращает CSV «название;план в заказах в день» в json для дневного отчёта.
// node scripts/make-orders-plan.js --csv input/plan-orders.csv --out input/plan-orders.json
const fs = require('fs');
const path = require('path');

function arg(name, def) {
  const i = process.argv.indexOf('--' + name);
  return i < 0 ? def : process.argv[i + 1];
}

const csv = arg('csv');
const out = arg('out', path.join(__dirname, '..', 'input', 'plan-orders.json'));
if (!csv) throw new Error('укажи --csv <файл>');

const text = fs.readFileSync(csv, 'utf8').replace(/^\uFEFF/, '');
const lines = text.split(/\r?\n/).filter((l) => l.trim());
const items = [];
for (const line of lines) {
  const parts = line.split(/[;\t]/).map((s) => s.trim());
  if (parts.length < 2) continue;
  const name = parts[0];
  const val = Number(parts[1].replace(',', '.'));
  if (!name || !Number.isFinite(val)) continue;
  items.push({ name, ordersPerDay: val });
}

fs.writeFileSync(out, JSON.stringify(items, null, 1), 'utf8');
console.log('готово:', out, '| строк:', items.length);
