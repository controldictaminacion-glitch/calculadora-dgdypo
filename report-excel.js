/* ExcelJS is vendored under vendor/ with its MIT license. */
function getExcelReportData() {
  return {
    startDate: document.getElementById('startDate').value,
    endDate: document.getElementById('endDate').value,
    days: calculatedDays,
    organization: getOrganizationValue(),
    funding: getFundingValue(),
    validatedBy: document.getElementById('validatedBy').value.trim(),
    generatedAt: new Date(),
    levels: levelsData.map((item, index) => ({ ...item, quantity: quantities[index] || 0 }))
      .filter(item => item.quantity > 0)
  };
}

function buildExcelWorkbook(data = getExcelReportData()) {
  if (typeof ExcelJS === 'undefined') throw Error('No se pudo cargar la biblioteca de Excel. Recargue la página.');
  if (!data.startDate || !data.endDate || data.days <= 0) throw Error('Seleccione un periodo válido antes de exportar.');
  const wb = new ExcelJS.Workbook();
  const generatedAt = new Date(data.generatedAt || Date.now());
  wb.creator = 'DGDYPO';
  wb.title = 'Honorarios por temporalidad';
  wb.subject = 'Proyección de honorarios SUN';
  wb.created = generatedAt;
  wb.modified = generatedAt;
  wb.calcProperties.fullCalcOnLoad = true;
  const sheet = wb.addWorksheet('Honorarios', {
    properties: { tabColor: { argb: 'FF691C32' }, defaultRowHeight: 22 },
    views: [{ state: 'frozen', ySplit: 12, xSplit: 1, topLeftCell: 'B13', showGridLines: false }],
    pageSetup: {
      paperSize: 1, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
      horizontalCentered: true, printTitlesRow: '1:12',
      margins: { left: 0.3, right: 0.3, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 }
    },
    headerFooter: { oddFooter: '&LDGDYPO · Honorarios SUN&R&P de &N' }
  });
  const color = { guinda: 'FF691C32', accent: 'FF9F2241', gold: 'FFBC955C', ink: 'FF1F2937', muted: 'FF64748B', light: 'FFF7F3EC', line: 'FFE2E8F0', white: 'FFFFFFFF' };
  const currency = '"$"#,##0.00;[Red]("$"#,##0.00)';
  sheet.columns = [{ width: 17 }, { width: 24 }, { width: 24 }, { width: 20 }, { width: 28 }, { width: 29 }];
  function styleRow(index, fill) {
    const row = sheet.getRow(index);
    for (let col = 1; col <= 6; col++) {
      const cell = row.getCell(col);
      cell.font = { name: 'Arial', size: 11, color: { argb: color.ink } };
      cell.alignment = { vertical: 'middle' };
      if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    }
    return row;
  }
  function mergeText(range, text, options = {}) {
    sheet.mergeCells(range);
    const cell = sheet.getCell(range.split(':')[0]);
    cell.value = text;
    cell.font = { name: 'Arial', size: 11, color: { argb: color.ink }, ...options.font };
    cell.alignment = { vertical: 'middle', wrapText: true, ...options.alignment };
    return cell;
  }
  for (let index = 1; index <= 12; index++) styleRow(index);
  sheet.getRow(1).height = 9;
  sheet.getRow(2).height = 30;
  mergeText('A2:F2', 'Honorarios por temporalidad', { font: { size: 17, bold: true, color: { argb: color.guinda } } });
  mergeText('A3:F3', 'DGDYPO · Honorarios SUN · Importes en pesos mexicanos (MXN)', { font: { size: 10, color: { argb: color.muted } } });
  for (let col = 1; col <= 6; col++) sheet.getCell(3, col).border = { bottom: { style: 'medium', color: { argb: color.gold } } };
  sheet.getRow(4).height = 9;
  function metadata(row, label, value) {
    sheet.getCell(row, 1).value = label;
    sheet.getCell(row, 1).font = { name: 'Arial', size: 11, bold: true, color: { argb: color.guinda } };
    mergeText(`B${row}:F${row}`, value || 'Sin especificar');
    sheet.getRow(row).height = Math.max(25, Math.ceil(String(value || '').length / 120) * 16 + 8);
  }
  metadata(5, 'Organismo', data.organization);
  metadata(6, 'Financiamiento', data.funding);
  const serialDate = value => (Date.parse(`${value}T00:00:00Z`) - Date.UTC(1899, 11, 30)) / 86400000;
  sheet.getRow(7).values = ['Fecha de inicio', serialDate(data.startDate), 'Fecha de fin', serialDate(data.endDate), 'Días computados', data.days];
  styleRow(7);
  for (const column of [1, 3, 5]) sheet.getCell(7, column).font = { name: 'Arial', size: 11, bold: true, color: { argb: color.guinda } };
  for (const column of [2, 4]) { sheet.getCell(7, column).numFmt = 'dd/mm/yyyy'; sheet.getCell(7, column).alignment = { horizontal: 'center', vertical: 'middle' }; }
  sheet.getCell('F7').numFmt = '#,##0';
  sheet.getCell('F7').alignment = { horizontal: 'right', vertical: 'middle' };
  metadata(8, 'Validó por', data.validatedBy);
  metadata(9, 'Elaboración', new Intl.DateTimeFormat('es-MX', { dateStyle: 'long', timeStyle: 'short', timeZone: 'America/Mexico_City' }).format(generatedAt) + ' (Ciudad de México)');
  sheet.getRow(10).height = 9;
  mergeText('A11:F11', 'Niveles seleccionados', { font: { bold: true, color: { argb: color.guinda } } });
  sheet.getRow(12).values = ['Nivel', 'Monto mensual base', 'Tarifa diaria\n(Base / 30)', 'Cantidad de folios', 'Costo diario\npor folios', 'Subtotal de la\ntemporalidad'];
  sheet.getRow(12).height = 36;
  for (let col = 1; col <= 6; col++) {
    const cell = sheet.getCell(12, col);
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: color.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color.guinda } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { right: { style: 'thin', color: { argb: color.white } } };
  }
  const levels = data.levels.filter(item => Number(item.quantity) > 0);
  let folios = 0, daily = 0, grand = 0, monthly = 0;
  levels.forEach((item, index) => {
    const rowNumber = 13 + index;
    const quantity = Number(item.quantity);
    const base = Number(item.base);
    const dailyRate = base / 30;
    const dailyCost = dailyRate * quantity;
    const subtotal = dailyCost * data.days;
    const row = sheet.getRow(rowNumber);
    row.values = [
      `Nivel ${item.level}`, base, { formula: `B${rowNumber}/30`, result: dailyRate }, quantity,
      { formula: `C${rowNumber}*D${rowNumber}`, result: dailyCost },
      { formula: `E${rowNumber}*$F$7`, result: subtotal }
    ];
    styleRow(rowNumber, index % 2 ? color.light : color.white);
    row.height = 26;
    for (let col = 1; col <= 6; col++) {
      const cell = row.getCell(col);
      cell.alignment = { vertical: 'middle', horizontal: col === 1 ? 'left' : 'right' };
      cell.border = { bottom: { style: 'hair', color: { argb: color.line } } };
      if ([2, 3, 5, 6].includes(col)) cell.numFmt = currency;
      if (col === 4) cell.numFmt = '#,##0';
      if (col === 1 || col === 6) cell.font = { ...cell.font, bold: true };
    }
    folios += quantity; daily += dailyCost; grand += subtotal; monthly += base * quantity;
  });
  const lastDataRow = 12 + levels.length;
  if (levels.length) sheet.autoFilter = `A12:F${lastDataRow}`;
  else { styleRow(13); mergeText('A13:F13', 'Sin niveles con folios asignados.', { font: { color: { argb: color.muted } } }); }
  const totalRow = Math.max(lastDataRow, 13) + 1;
  styleRow(totalRow, color.light);
  mergeText(`A${totalRow}:C${totalRow}`, 'TOTALES', { font: { bold: true, color: { argb: color.guinda } } });
  for (const [col, cached] of [['D', folios], ['E', daily], ['F', grand]]) {
    const cell = sheet.getCell(`${col}${totalRow}`);
    cell.value = levels.length ? { formula: `SUM(${col}13:${col}${lastDataRow})`, result: cached } : 0;
    cell.numFmt = col === 'D' ? '#,##0' : currency;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: color.guinda } };
    cell.alignment = { horizontal: 'right', vertical: 'middle' };
  }
  sheet.getRow(totalRow).height = 30;
  for (let col = 1; col <= 6; col++) sheet.getCell(totalRow, col).border = { top: { style: 'medium', color: { argb: color.gold } } };
  const monthlyRow = totalRow + 2;
  styleRow(monthlyRow);
  mergeText(`A${monthlyRow}:E${monthlyRow}`, 'Base mensual de los folios seleccionados (30 días)', { font: { color: { argb: color.muted } }, alignment: { horizontal: 'right' } });
  sheet.getCell(`F${monthlyRow}`).value = levels.length ? { formula: `SUMPRODUCT(B13:B${lastDataRow},D13:D${lastDataRow})`, result: monthly } : 0;
  sheet.getCell(`F${monthlyRow}`).numFmt = currency;
  sheet.getCell(`F${monthlyRow}`).alignment = { horizontal: 'right', vertical: 'middle' };
  const ruleRow = monthlyRow + 2;
  styleRow(ruleRow);
  mergeText(`A${ruleRow}:F${ruleRow}`, 'Cómputo de temporalidad: enero a noviembre, meses comerciales de 30 días; diciembre, días reales.', { font: { size: 10, color: { argb: color.muted } } });
  sheet.pageSetup.printArea = `A1:F${ruleRow}`;
  return wb;
}

async function exportExcel() {
  const button = document.getElementById('exportExcelButton');
  if (button?.disabled) return;
  if (button) { button.disabled = true; button.setAttribute('aria-busy', 'true'); }
  let objectURL;
  try {
    const data = getExcelReportData();
    const workbook = buildExcelWorkbook(data);
    const bytes = await workbook.xlsx.writeBuffer();
    objectURL = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a');
    link.href = objectURL;
    link.download = `Honorarios_DGDYPO_${data.startDate}_a_${data.endDate}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    console.error(error);
    alert(error.message || 'No se pudo generar el archivo Excel. Intente de nuevo.');
  } finally {
    if (objectURL) setTimeout(() => URL.revokeObjectURL(objectURL), 1000);
    if (button) { button.disabled = false; button.removeAttribute('aria-busy'); }
  }
}
