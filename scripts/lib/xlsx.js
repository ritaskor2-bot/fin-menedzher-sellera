// Общий модуль записи xlsx без внешних зависимостей, кроме jszip.
const fs = require('fs');
const path = require('path');
const JSZip = (() => {
  const candidates = [
    path.join(__dirname, '..', '..', '..', 'tmp', 'node_modules', 'jszip'),
    path.join(__dirname, '..', '..', 'tmp', 'node_modules', 'jszip'),
  ];
  for (const c of candidates) {
    try { return require(c); } catch (e) { /* next */ }
  }
  throw new Error('не найден jszip. Установи: npm i jszip в корне проекта');
})();

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const colName = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts>
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF666666"/><name val="Calibri"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills>
<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="10"><xf xfId="0"/>
<xf xfId="0" fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment wrapText="1" vertical="center" horizontal="center"/></xf>
<xf xfId="0" numFmtId="164" applyNumberFormat="1"/>
<xf xfId="0" numFmtId="165" applyNumberFormat="1"/>
<xf xfId="0" fontId="1" numFmtId="164" applyNumberFormat="1" applyFont="1"/>
<xf xfId="0" fontId="1" applyFont="1"/>
<xf xfId="0" fontId="2" applyFont="1"/>
<xf xfId="0" fontId="3" applyFont="1"/>
<xf xfId="0" fontId="1" fillId="4" applyFont="1" applyFill="1" numFmtId="164" applyNumberFormat="1"/>
<xf xfId="0" fontId="1" fillId="3" applyFont="1" applyFill="1"/>
</cellXfs></styleSheet>`;

function sheetXml(rows, { widths = [], freeze = null, freezeCols = 0 } = {}) {
  const cols = widths.length ? `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
  const body = rows.map((row, ri) => {
    const cells = row.map((cell, ci) => {
      if (cell === null || cell === undefined || cell === '') return '';
      const ref = colName(ci + 1) + (ri + 1);
      if (typeof cell === 'object') {
        const s = cell.s === undefined ? '' : ` s="${cell.s}"`;
        if (cell.f) return `<c r="${ref}"${s}><f>${esc(cell.f)}</f></c>`;
        if (typeof cell.v === 'number') return `<c r="${ref}"${s}><v>${cell.v}</v></c>`;
        return `<c r="${ref}"${s} t="inlineStr"><is><t>${esc(cell.v)}</t></is></c>`;
      }
      if (typeof cell === 'number') return `<c r="${ref}"><v>${cell}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t>${esc(cell)}</t></is></c>`;
    }).join('');
    return `<row r="${ri + 1}">${cells}</row>`;
  }).join('');
  const pane = freeze ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${freeze}" xSplit="${freezeCols}" topLeftCell="${colName(freezeCols + 1)}${freeze + 1}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pane}${cols}<sheetData>${body}</sheetData></worksheet>`;
}

async function writeXlsx(file, sheets) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', stylesXml);
  sheets.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows, s.opts || {})));
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

module.exports = { writeXlsx, colName, esc };
