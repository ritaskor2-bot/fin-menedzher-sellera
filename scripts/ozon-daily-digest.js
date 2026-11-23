// Дневной дайджест по Ozon: заказы против плана, топ-17, рекомендации.
// Запуск:
// node scripts/ozon-daily-digest.js --file input/optiflou-rnp-2026-12-01-2026-12-01.xlsx [--plan input/plan-ozon-orders.json] [--send]
// План в json: [{ "name": "Артикул", "ordersPerDay": 12.5 }, ...]
const fs = require('fs');
const path = require('path');
const https = require('https');
const { colName } = require('./lib/xlsx');
const JSZip = (() => {
  const c = [path.join(__dirname, '..', '..', 'tmp', 'node_modules', 'jszip')];
  for (const x of c) { try { return require(x); } catch (e) {} }
  throw new Error('нужен jszip: npm i jszip');
})();

const dec = (s) => String(s || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const num = (v) => Math.round(v || 0).toLocaleString('ru-RU');
const pct = (v) => (Number.isFinite(v) ? (v * 100).toFixed(1).replace('.', ',') : '—') + '%';
const p100 = (v) => (Number.isFinite(v) ? v.toFixed(1).replace('.', ',') : '—') + '%';
const norm = (s) => String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]/g, '');

function arg(name, def) {
  const i = process.argv.indexOf('--' + name);
  if (i < 0) return def;
  const v = process.argv[i + 1];
  return v === undefined ? true : v;
}

async function readSheet(file, sheetName) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const entries = [...wb.matchAll(/<sheet[^>]*\/>/g)].map((m) => m[0]);
  const entry = entries.find((s) => s.includes(`name="${sheetName}"`)) || entries[0];
  const rid = entry.match(/r:id="([^"]+)"/)[1];
  const rel = [...rels.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].find((m) => m[1] === rid);
  const xml = await zip.file('xl/' + rel[2].replace(/^\/?xl\//, '')).async('string');
  const ss = [];
  if (zip.files['xl/sharedStrings.xml']) {
    const x = await zip.file('xl/sharedStrings.xml').async('string');
    for (const m of x.matchAll(/<si>([\s\S]*?)<\/si>/g)) ss.push([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => dec(t[1])).join(''));
  }
  const rows = [];
  for (const r of xml.matchAll(/<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const cells = [];
    for (const c of r[2].matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const idx = [...c[1]].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const t = (c[3].match(/t="([^"]+)"/) || [])[1];
      const v = (c[4] || '').match(/<v>([\s\S]*?)<\/v>/);
      const is = (c[4] || '').match(/<t[^>]*>([\s\S]*?)<\/t>/);
      cells[idx] = t === 's' ? ss[+v[1]] : t === 'inlineStr' || t === 'str' ? dec(is ? is[1] : '') : v ? Number(v[1]) : null;
    }
    rows.push({ r: Number(r[1]), cells });
  }
  return rows;
}

(async () => {
  const file = arg('file');
  if (!file) throw new Error('укажи --file <путь к дневному отчёту Ozon>');
  const planFile = arg('plan');
  const plan = planFile && planFile !== true ? JSON.parse(fs.readFileSync(planFile, 'utf8')) : [];
  const planBy = new Map((Array.isArray(plan) ? plan : []).map((p) => [norm(p.name), p]));

  const rows = await readSheet(file, 'По товарам');
  const items = [];
  let orders = 0;
  for (const row of rows.slice(1)) {
    if (row.cells[0] !== 'Артикул') continue;
    const o = {
      name: String(row.cells[1] || ''),
      orders: row.cells[9] || 0,
      ads: row.cells[15] || 0,
      price: row.cells[25] || 0,
      spp: row.cells[26] || 0,
      disc: row.cells[29] || 0,
      shows: row.cells[30] || 0,
      clicks: row.cells[31] || 0,
      carts: row.cells[32] || 0,
      crCart: row.cells[33] || 0,
      crOrder: row.cells[34] || 0,
      ctr: row.cells[35] || 0,
      drr: row.cells[54] || 0,
    };
    items.push(o);
    orders += o.orders;
  }

  const planOrders = [...planBy.values()].reduce((s, p) => s + (p.ordersPerDay || 0), 0);
  const top = [...items].sort((a, b) => b.orders - a.orders).slice(0, 17);

  const dateMatch = path.basename(file).match(/(\d{4}-\d{2}-\d{2})/);
  const date = dateMatch ? dateMatch[1].split('-').reverse().join('.') : '';

  const lines = [];
  lines.push(`📅 Ozon · ${date}`);
  lines.push(`Заказы: ${num(orders)}${planOrders ? ` из плана ${num(planOrders)}/день — ${pct(orders / planOrders)}` : ''}`);
  lines.push('');
  top.forEach((x, i) => {
    const p = planBy.get(norm(x.name));
    const planDay = p ? p.ordersPerDay || 0 : 0;
    const exec = planDay > 0 ? x.orders / planDay : null;
    const actions = [];
    if (exec !== null && exec < 0.5) actions.push(`план ${pct(exec)} — мало заказов`);
    if (exec !== null && exec > 1.5) actions.push(`план ${pct(exec)} — проверить остаток`);
    if (x.drr > 15) actions.push(`ДРР ${p100(x.drr)} — снизить ставки`);
    if (x.ctr < 5) actions.push(`CTR ${p100(x.ctr)} — фото и цена`);
    if (x.crCart < 3) actions.push(`в корзину ${p100(x.crCart)} — цена выше ожиданий`);
    if (x.crOrder < 12) actions.push(`корзина→заказ ${p100(x.crOrder)} — цена и доставка`);
    lines.push(`${i + 1}) ${x.name}`);
    lines.push(`Заказы ${num(x.orders)}${planDay ? ` из ${num(planDay)}/день (${pct(exec)})` : ''} · СПП ${p100(x.spp)}, скидка ${p100(x.disc)}`);
    lines.push(`ДРР ${p100(x.drr)} · показы ${num(x.shows)}, CTR ${p100(x.ctr)}, корзина ${p100(x.crCart)}, корзина→заказ ${p100(x.crOrder)}`);
    lines.push(`Делать: ${actions.slice(0, 2).join('; ') || 'показатели в норме'}`);
    lines.push('');
  });

  const text = lines.join('\n');
  console.log(text);

  if (process.argv.includes('--send')) {
    const token = process.env.TG_TOKEN;
    const chat = process.env.TG_CHAT;
    if (!token || !chat) throw new Error('для отправки задай TG_TOKEN и TG_CHAT');
    const body = JSON.stringify({ chat_id: chat, text: text.slice(0, 4000) });
    const res = await new Promise((resolve, reject) => {
      const req = https.request({
        hostname: 'api.telegram.org',
        path: `/bot${token}/sendMessage`,
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
      }, (r) => { let d = ''; r.on('data', (c) => (d += c)); r.on('end', () => resolve(JSON.parse(d))); });
      req.on('error', reject);
      req.write(body);
      req.end();
    });
    console.log('отправка:', res.ok ? 'ок' : JSON.stringify(res));
  }
})();
