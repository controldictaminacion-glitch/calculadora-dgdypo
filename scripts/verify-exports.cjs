// Focused checks for exported report data. No browser or installed Office needed.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'calculadora_de_honorarios_por_temporalidad_cdmx.html'), 'utf8');
const elements = new Map();
class Element {
  constructor(tagName = 'div') {
    this.tagName = tagName;
    this.value = '';
    this.textContent = '';
    this.innerHTML = '';
    this.dataset = {};
    this.children = [];
    this.classList = { add() {}, remove() {}, toggle() {} };
  }
  set id(value) { this._id = value; elements.set(value, this); }
  get id() { return this._id; }
  setAttribute() {}
  appendChild(child) { this.children.push(child); return child; }
  querySelector() { return null; }
  remove() {}
  click() {}
}
const document = {
  body: new Element('body'),
  getElementById(id) { return elements.get(id) || null; },
  createElement(tag) { return new Element(tag); },
  querySelector() { return null; },
};
for (const match of source.matchAll(/\bid="([^"]+)"/g)) {
  if (!elements.has(match[1])) { const element = new Element(); element.id = match[1]; }
}
const context = vm.createContext({
  document, console, Intl, Date, Blob, URL, setTimeout, clearTimeout, setImmediate, clearImmediate, TextEncoder, TextDecoder,
  tailwind: {},
});
context.window = context;
context.addEventListener = () => {};
context.print = () => {};
for (const match of source.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)) {
  const src = match[1].match(/\bsrc\s*=\s*["']([^"']+)["']/)?.[1];
  if (!src) new vm.Script(match[2]).runInContext(context);
  else if (!/^https?:/.test(src)) {
    const localSource = path.join(root, src.split('?')[0]);
    if (/exceljs/.test(src)) new vm.Script(fs.readFileSync(localSource, 'utf8'), { filename: localSource }).runInContext(context);
    else if (/jspdf\.umd/.test(src)) context.jspdf = require(localSource);
    else if (/autotable/.test(src)) require(localSource).applyPlugin(context.jspdf.jsPDF);
    else new vm.Script(fs.readFileSync(localSource, 'utf8'), { filename: localSource }).runInContext(context);
  }
}
const evaluate = code => vm.runInContext(code, context);
const get = id => document.getElementById(id);
const levels = JSON.parse(evaluate('JSON.stringify(levelsData)'));
levels.forEach((_, index) => {
  for (const prefix of ['qty', 'dailySubtotal', 'subtotal']) {
    const element = new Element(); element.id = `${prefix}-${index}`;
  }
});
const numberMXN = value => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
const selected = [3, 0, 0, 2, 0, 0, 0, 0, 0, 0, 5, 0];
const expected = quantities => quantities.reduce((total, qty, index) => ({
  folios: total.folios + qty,
  monthly: total.monthly + qty * levels[index].base,
  daily: total.daily + qty * levels[index].base / 30,
}), { folios: 0, monthly: 0, daily: 0 });

function setup(quantities, options = {}) {
  get('startDate').value = options.start || '2026-01-01';
  get('endDate').value = options.end || '2026-12-29';
  get('organization').value = options.organization || 'SECRETARÍA DE ADMINISTRACIÓN Y FINANZAS';
  get('organizationFallback').value = options.fallback || '';
  get('fundingType').value = options.funding || 'FISCALES';
  get('fundingOther').value = options.other || '';
  get('validatedBy').value = options.validator || 'MUESTRA DE VALIDACIÓN';
  evaluate(`quantities = ${JSON.stringify(quantities)}; showZeroRows = true; calculateDays();`);
}

function checkPdf(quantities, options = {}) {
  setup(quantities, options);
  evaluate('preparePrintReport()');
  const html = get('printReport').innerHTML;
  const tbody = html.match(/<tbody>([\s\S]*?)<\/tbody>/)[1];
  const levelNames = Array.from(tbody.matchAll(/<tr><td>(\d+)<\/td>/g), match => match[1]);
  assert.deepEqual(levelNames, levels.filter((_, index) => quantities[index] > 0).map(item => item.level));
  const totals = expected(quantities);
  const days = evaluate('calculatedDays');
  assert.ok(html.includes(numberMXN(totals.daily * days)), 'PDF grand total');
  assert.ok(html.includes(numberMXN(totals.monthly)), 'PDF monthly base');
  assert.ok(html.includes(numberMXN(totals.daily)), 'PDF daily cost');
  if (totals.folios === 0) assert.ok(tbody.includes('Sin niveles con folios asignados.'));
  const pdf = evaluate('buildPDFReport()');
  assert.equal(pdf.getNumberOfPages(), 1, 'downloaded PDF must fit one page');
  const pdfRows = pdf.lastAutoTable.body;
  if (totals.folios) {
    assert.deepEqual(Array.from(pdfRows, row => row.raw[0]), levels.filter((_, index) => quantities[index] > 0).map(item => item.level));
    assert.equal(pdf.lastAutoTable.foot[0].raw[1], String(totals.folios));
    assert.equal(pdf.lastAutoTable.foot[0].raw[3], numberMXN(totals.daily * days));
  }
  assert.ok(Math.abs(pdf.internal.pageSize.getWidth() - 215.9) < .01, 'Letter width');
  assert.ok(Math.abs(pdf.internal.pageSize.getHeight() - 279.4) < .01, 'Letter height');
  // A one-page file can still have clipped text: check the emitted text baselines.
  const commands = pdf.internal.pages[1].join('\n');
  for (const match of commands.matchAll(/([-\d.]+) ([-\d.]+) Td/g)) {
    const baseline = Number(match[2]);
    assert.ok(baseline >= 25 && baseline <= 792, `PDF text baseline inside Letter page: ${baseline}`);
  }
  return { html, totals, days, pdf };
}

function previewJson(sheet) {
  const rows = [];
  const cells = [];
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    rows.push({ number: rowNumber, height: row.height || sheet.properties.defaultRowHeight || 15 });
    for (let col = 1; col <= sheet.columnCount; col++) {
      const cell = row.getCell(col);
      const raw = cell.value;
      const value = raw && typeof raw === 'object' && 'formula' in raw ? raw.result : raw;
      let text = value == null ? '' : String(value);
      if (value instanceof Date) text = new Intl.DateTimeFormat('es-MX', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC' }).format(value);
      else if (typeof value === 'number') {
        if (/d.*m.*y/i.test(cell.numFmt)) text = new Intl.DateTimeFormat('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' }).format(new Date(Date.UTC(1899, 11, 30) + value * 86400000));
        else if (/\$/.test(cell.numFmt)) text = numberMXN(value);
        else text = new Intl.NumberFormat('es-MX', { maximumFractionDigits: /0\.00/.test(cell.numFmt) ? 2 : 0 }).format(value);
      }
      cells.push({
        address: cell.address, row: rowNumber, col, value: raw, text,
        numFmt: cell.numFmt, font: cell.font, fill: cell.fill, border: cell.border,
        alignment: cell.alignment, isMerged: cell.isMerged,
        master: cell.isMerged ? cell.master.address : null,
      });
    }
  }
  return {
    sheetName: sheet.name, rowCount: sheet.rowCount, columnCount: sheet.columnCount,
    columns: Array.from({ length: sheet.columnCount }, (_, index) => {
      const column = sheet.getColumn(index + 1);
      return { number: index + 1, letter: column.letter, width: column.width || 8.43 };
    }),
    rows, cells, merges: sheet.model.merges,
  };
}

function checkWorkbook(sheet, quantities) {
  let header;
  let total;
  sheet.eachRow(row => {
    if (row.getCell(1).text.trim() === 'Nivel') header = row.number;
    if (/^TOTALES?$/i.test(row.getCell(1).text.trim())) total = row.number;
  });
  assert.ok(header, 'Excel level table header');
  assert.ok(total > header, 'Excel totals below header');
  const numeric = cell => cell.value && typeof cell.value === 'object' && 'formula' in cell.value ? cell.value.result : cell.value;
  const data = [];
  for (let rowNumber = header + 1; rowNumber < total; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const level = row.getCell(1).text.replace(/^Nivel\s+/, '');
    if (/^\d{4}$/.test(level)) data.push({ level, row });
  }
  assert.deepEqual(data.map(item => item.level), levels.filter((_, index) => quantities[index] > 0).map(item => item.level));
  for (const { level, row } of data) {
    const index = levels.findIndex(item => item.level === level);
    const rowIndex = data.findIndex(item => item.level === level);
    for (let column = 1; column <= 6; column++) {
      const cell = row.getCell(column);
      assert.equal(cell.font.name, 'Arial', 'consistent table font');
      assert.equal(cell.font.size, 11, 'legible table font size');
      assert.equal(cell.fill.pattern, 'solid', 'solid table row fill');
      assert.equal(cell.fill.fgColor.argb, rowIndex % 2 ? 'FFF7F3EC' : 'FFFFFFFF', 'alternating table row shading');
    }
    assert.equal(numeric(row.getCell(2)), levels[index].base, 'numeric monthly base');
    assert.equal(numeric(row.getCell(4)), quantities[index], 'numeric folio quantity');
    for (const column of [3, 5, 6]) {
      assert.equal(typeof row.getCell(column).value.formula, 'string', 'Excel calculation formula');
      assert.equal(typeof row.getCell(column).value.result, 'number', 'Excel cached calculation result');
      assert.ok(/\$/.test(row.getCell(column).numFmt), 'currency number format');
    }
    const daily = levels[index].base / 30;
    assert.ok(Math.abs(numeric(row.getCell(3)) - daily) < 1e-7, 'daily rate retains precision');
    assert.ok(Math.abs(numeric(row.getCell(5)) - daily * quantities[index]) < 1e-7, 'daily amount');
    assert.ok(Math.abs(numeric(row.getCell(6)) - daily * quantities[index] * evaluate('calculatedDays')) < 1e-7, 'season amount');
  }
  const totals = expected(quantities);
  assert.equal(numeric(sheet.getCell(total, 4)), totals.folios, 'total folios in quantity column');
  assert.ok(Math.abs(numeric(sheet.getCell(total, 5)) - totals.daily) < 1e-7, 'total daily cost');
  assert.ok(Math.abs(numeric(sheet.getCell(total, 6)) - totals.daily * evaluate('calculatedDays')) < 1e-7, 'total season cost');
  assert.ok(sheet.views.some(view => view.state === 'frozen' && view.ySplit >= header), 'frozen table header');
  if (totals.folios) assert.ok(sheet.autoFilter, 'filterable table');
  assert.equal(sheet.pageSetup.fitToWidth, 1);
  assert.equal(sheet.pageSetup.fitToHeight, 1);
  sheet.eachRow(row => row.eachCell(cell => assert.ok(!cell.value?.error, `no formula error in ${cell.address}`)));
  return { header, total, totals };
}

async function checkExcel(quantities, options = {}) {
  setup(quantities, options);
  const workbook = evaluate('buildExcelWorkbook(getExcelReportData())');
  const sheet = workbook.worksheets[0];
  checkWorkbook(sheet, quantities);
  const buffer = await workbook.xlsx.writeBuffer();
  const reopened = new context.ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);
  checkWorkbook(reopened.worksheets[0], quantities);
  return { workbook: reopened, sheet: reopened.worksheets[0], buffer: Buffer.from(buffer) };
}

async function checkPDFFileNames(samples, output) {
  const fixture = fs.readFileSync(path.join(__dirname, 'fixtures', 'organizations.csv'), 'utf8');
  const parsed = JSON.parse(evaluate(`JSON.stringify(parseCSVRows(${JSON.stringify('\ufeff' + fixture.replace(/\n/g, '\r\n'))}))`));
  assert.deepEqual(parsed[2], ['Organismo, "de prueba"', 'ODP']);
  assert.deepEqual(parsed[3], ['Organismo\r\nmultilínea', 'OML']);
  const previousFetch = context.fetch;
  context.fetch = async url => {
    assert.ok(url.endsWith('&range=B2:C'), 'fetch organization and acronym from the same sheet rows');
    return { ok: true, text: async () => fixture };
  };
  try {
    await evaluate('loadOrganizations()');
    assert.equal(evaluate('organizationAcronyms.size'), 6, 'blank names ignored and duplicate names combined');
    assert.equal(evaluate('organizationAcronyms.get(normalizeOrganizationName(\'Organismo, "de prueba"\'))'), 'ODP');
    assert.equal(evaluate('organizationAcronyms.get(normalizeOrganizationName("Organismo multilínea"))'), 'OML');
    assert.equal(evaluate('organizationAcronyms.get(normalizeOrganizationName("Duplicado de prueba"))'), 'DUP', 'blank duplicate cannot erase acronym');
    const jefatura = 'Jefatura de Gobierno de la Ciudad de México';
    setup(selected, { organization: jefatura });
    const name = evaluate('getPDFFileName()');
    assert.equal(name, 'Honorarios_JGCDMX_10_Folios_FISCALES_$5,365,853.33_del_2026-01-01_al_2026-12-29.pdf');
    if (samples) fs.writeFileSync(path.join(output, name), Buffer.from(evaluate('buildPDFReport().output("arraybuffer")')));
    get('organization').value = 'Agencia Digital de Innovación Pública de la Ciudad de México';
    assert.ok(evaluate('getPDFFileName()').startsWith('Honorarios_ADIP_'), 'switching organization changes acronym');
    setup(selected, { organization: jefatura, start: '2026-03-15', end: '2026-04-14' });
    assert.equal(evaluate('getPDFFileName()'), 'Honorarios_JGCDMX_10_Folios_FISCALES_$448,400.00_del_2026-03-15_al_2026-04-14.pdf');
    setup(levels.map(() => 0), { organization: jefatura });
    assert.equal(evaluate('getPDFFileName()'), 'Honorarios_JGCDMX_0_Folios_FISCALES_$0.00_del_2026-01-01_al_2026-12-29.pdf');
    setup(selected, { organization: jefatura, funding: 'PROPIOS' });
    assert.ok(evaluate('getPDFFileName()').includes('_10_Folios_PROPIOS_$5,365,853.33_'));
    setup(selected, { organization: jefatura, funding: 'OTRO', other: 'ÁREAS / "PROPIAS": <>?*' });
    assert.ok(evaluate('getPDFFileName()').includes('_10_Folios_AREAS_PROPIAS_$5,365,853.33_'), 'custom financing text is sanitized');
    get('fundingType').value = '';
    assert.ok(evaluate('getPDFFileName()').includes('_10_Folios_Sin_financiamiento_'));
    const increased = [...selected]; increased[0]++;
    setup(increased, { organization: jefatura, funding: 'PROPIOS' });
    assert.equal(evaluate('getPDFFileName()'), 'Honorarios_JGCDMX_11_Folios_PROPIOS_$6,374,643.33_del_2026-01-01_al_2026-12-29.pdf');
    setup(selected, { fallback: 'ÁREA / "DE PRUEBA": <>?* | RECURSOS' });
    const fallback = evaluate('getPDFFileName()');
    assert.ok(fallback.startsWith('Honorarios_AREA_DE_PRUEBA_RECURSOS_10_Folios_FISCALES_'), 'manual name uses sanitized full name');
    assert.ok(!/[<>:"/\\|?*\s]/.test(fallback), 'filename has no forbidden path characters');
    get('organizationFallback').value = '';
    get('organization').value = 'Organismo sin acrónimo';
    assert.ok(evaluate('getPDFFileName()').startsWith('Honorarios_Organismo_sin_acronimo_'), 'missing acronym uses full name');
    console.log('PDF filename: sheet mapping, acronym, folio count, financing, currency, dates and safe fallbacks OK');
    if (samples) console.log(`Named PDF sample: ${name}`);
  } finally {
    context.fetch = previousFetch;
    setup(selected);
  }
}

async function main() {
  const samples = process.argv.includes('--samples');
  const output = path.join(root, 'output', 'previos');
  const temp = path.join(root, 'tmp', 'pdfs');
  if (samples) { fs.mkdirSync(output, { recursive: true }); fs.mkdirSync(temp, { recursive: true }); }
  const cases = [
    ['none', levels.map(() => 0)],
    ['selected', selected],
    ['all', levels.map(() => 1)],
  ];
  for (const [label, quantities] of cases) {
    const check = checkPdf(quantities);
    if (samples) fs.writeFileSync(path.join(label === 'selected' ? output : temp, label === 'selected' ? 'Honorarios_DGDYPO_ejemplo.pdf' : `honorarios-${label}.pdf`), Buffer.from(check.pdf.output('arraybuffer')));
    console.log(`PDF ${label}: ${quantities.filter(Boolean).length} levels; ${check.totals.folios} folios; ${check.days} days; ${numberMXN(check.totals.daily * check.days)}`);
  }
  const odd = checkPdf(selected, {
    start: '2026-03-15', end: '2026-04-14',
    fallback: 'ORGANISMO DE PRUEBA CON NOMBRE EXTENSO & <script>alert("x")</script> "ÁÉÍÓÚ"',
    funding: 'OTRO', other: 'recurso & <especial>', validator: 'Nombre <b>Prueba</b> & O\'Brien',
  });
  assert.equal(odd.days, 30);
  assert.ok(odd.html.includes('&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;'));
  assert.ok(odd.html.includes('RECURSO &amp; &lt;ESPECIAL&gt;'));
  assert.ok(odd.html.includes('Nombre &lt;b&gt;Prueba&lt;/b&gt; &amp; O&#39;Brien'));
  assert.ok(!odd.html.includes('<script>'));
  console.log('PDF custom dates, long fallback, OTRO and escaped characters: OK');

  const long = checkPdf(levels.map(() => 1), {
    fallback: 'DIRECCIÓN GENERAL DE DESARROLLO Y PLANEACIÓN ORGANIZACIONAL DE LA SECRETARÍA DE ADMINISTRACIÓN Y FINANZAS DEL GOBIERNO DE LA CIUDAD DE MÉXICO, COORDINACIÓN GENERAL DE ADMINISTRACIÓN Y DESARROLLO DE PERSONAL',
    funding: 'OTRO', other: 'PARTICIPACIONES EN INGRESOS FEDERALES Y RECURSOS PROPIOS PARA EL FORTALECIMIENTO DE LA ADMINISTRACIÓN PÚBLICA',
    validator: 'LICENCIADA MARÍA DE LOS ÁNGELES HERNÁNDEZ RODRÍGUEZ, DIRECTORA GENERAL DE ADMINISTRACIÓN',
  });
  if (samples) fs.writeFileSync(path.join(temp, 'honorarios-long.pdf'), Buffer.from(long.pdf.output('arraybuffer')));
  console.log('PDF all 12 levels and long metadata: one Letter page; text stays inside page');

  if (evaluate('typeof buildExcelWorkbook') !== 'function') {
    console.log('Excel builder not yet available. PDF checks completed.');
    return;
  }
  const tempExcel = path.join(root, 'tmp', 'excel');
  if (samples) fs.mkdirSync(tempExcel, { recursive: true });
  for (const [label, quantities] of cases) {
    const check = await checkExcel(quantities);
    if (samples) {
      fs.writeFileSync(path.join(label === 'selected' ? output : tempExcel, label === 'selected' ? 'Honorarios_DGDYPO_ejemplo.xlsx' : `honorarios-${label}.xlsx`), check.buffer);
      if (label === 'selected') fs.writeFileSync(path.join(root, 'tmp', 'preview-workbook.json'), JSON.stringify(previewJson(check.sheet), null, 2));
    }
    console.log(`Excel ${label}: selected levels, numeric values, formulas/caches, totals, styles and saved-file round trip OK`);
  }
  const special = await checkExcel(selected, {
    start: '2026-03-15', end: '2026-04-14',
    fallback: '=ORGANISMO & <PRUEBA> "ÁÉÍÓÚ" CON NOMBRE EXTENSO PARA VALIDAR LA PRESENTACIÓN',
    funding: 'OTRO', other: 'recurso & <especial>', validator: 'Nombre <b>Prueba</b> & O\'Brien',
  });
  const values = [];
  special.sheet.eachRow(row => row.eachCell(cell => { if (cell.type !== context.ExcelJS.ValueType.Merge) values.push(cell.value); }));
  assert.ok(values.some(value => typeof value === 'string' && value.includes('=ORGANISMO & <PRUEBA>')), 'organization stored as plain text');
  assert.ok(values.some(value => typeof value === 'string' && value.includes('RECURSO & <ESPECIAL>')), 'OTRO funding preserved');
  assert.ok(values.some(value => typeof value === 'string' && value.includes('Nombre <b>Prueba</b>')), 'validator preserved as plain text');
  console.log('Excel custom dates, fallback, OTRO and special characters: OK');
  await checkPDFFileNames(samples, output);
}
main().catch(error => { console.error(error); process.exitCode = 1; });
