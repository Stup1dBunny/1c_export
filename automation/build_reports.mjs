import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const root = path.dirname(fileURLToPath(import.meta.url));
const manifestPath = path.resolve(process.argv[2]);
const dataRoot = path.dirname(manifestPath);
const manifest = JSON.parse(await fs.readFile(manifestPath, 'utf8'));
const outputDir = manifest.output_dir ? path.resolve(dataRoot, manifest.output_dir) : path.resolve(root, '../outputs/reports');
await fs.mkdir(outputDir, { recursive: true });
const schemas = {
  'Продажи': { headers: ['Дата отчёта', 'Адрес магазина', 'Выручка, ₽', 'Чеки', 'Алкоголь, ед.', 'Количество возвратов', 'Сумма возвратов, ₽', 'Средний чек, ₽'],
    keys: ['report_date', 'shop', 'revenue', 'receipt_count', 'alcohol_quantity', 'refund_receipt_count', 'refund_amount', 'average_receipt'],
    widths: [16, 62, 19, 13, 18, 18, 20, 20], numeric: [2,3,4,5,6,7], table: 'Sales' },
  'Остатки': { headers: ['Дата отчёта', 'Адрес магазина', 'Номенклатура', 'Поставщик', 'Начальный остаток', 'Приход', 'Расход', 'Конечный остаток'],
    keys: ['report_date', 'shop', 'product', 'supplier', 'opening', 'incoming', 'outgoing', 'closing'], widths: [16, 48, 62, 30, 17, 14, 14, 17], numeric: [4,5,6,7], table: 'Stock' },
  'Бонусы': { headers: ['Дата покупки', 'Дисконтная карта — номер и владелец', 'Сумма по ДК, ₽', 'Количество чеков', 'Магазин'],
    keys: ['date', 'card', 'amount', 'receipt_count', 'shop'], widths: [16, 49, 20, 16, 58], numeric: [2,3], table: 'Bonuses' },
  'Контроль': { headers: ['Состояние', 'Магазин', 'Причина', 'Исходные значения', 'Обнаружено', 'Проверено'],
    keys: ['state', 'shop', 'message', 'detail', 'first_seen', 'last_checked'], widths: [16, 54, 65, 75, 26, 26], numeric: [], table: 'Control' },
};
function col(n) { return String.fromCharCode(65 + n); }
for (const item of manifest.reports) {
  let spec = schemas[item.type];
  if (!spec || !/^\d{4}-\d{2}-\d{2}$/.test(item.date)) throw Error('Invalid report metadata');
  if (item.end_date && (!/^\d{4}-\d{2}-\d{2}$/.test(item.end_date) || item.end_date < item.date)) throw Error('Invalid report period');
  const displayDate = item.date.split('-').reverse().join('.') + (item.end_date ? '-' + item.end_date.split('-').reverse().join('.') : '');
  if (item.end_date && ['Продажи', 'Остатки'].includes(item.type)) {
    spec = { ...spec, headers: ['Период отчёта', ...spec.headers.slice(1)], widths: [28, ...spec.widths.slice(1)] };
  }
  const source = JSON.parse(await fs.readFile(path.resolve(dataRoot, item.data), 'utf8'));
  const preliminary = source.quality === 'preliminary';
  if (preliminary) spec = { ...spec, headers: [...spec.headers, 'Статус сверки'], keys: [...spec.keys, 'verification'], widths: [...spec.widths, 24] };
  if (item.type === 'Бонусы' && source.card_source_verified !== true) throw Error('Bulk card source has not been verified');
  const rows = source.rows.map(r => spec.keys.map((key, i) => {
    if (key === 'report_date') return item.end_date ? displayDate : new Date(item.date + 'T00:00:00Z');
    if (key === 'average_receipt') {
      if (['revenue', 'receipt_count', 'refund_amount', 'refund_receipt_count'].some(k => r[k] === null)) return null;
      const count = Number(r.receipt_count) - Number(r.refund_receipt_count);
      if (count < 0) throw Error('Refund check count exceeds all checks');
      return count ? (Number(r.revenue) + Number(r.refund_amount)) / count : null;
    }
    if (!(key in r)) throw Error(`Missing field ${key}`);
    if (item.type === 'Контроль' && ['first_seen', 'last_checked'].includes(key)) return new Date(r[key].slice(0,19) + 'Z');
    if (item.type === 'Бонусы' && key === 'date') {
      if (r[key] < item.date || r[key] > (item.end_date || item.date)) throw Error('Bonus row date differs from report period');
      return new Date(r[key] + 'T00:00:00Z');
    }
    if (!spec.numeric.includes(i)) return String(r[key]).startsWith('=') ? "'" + r[key] : String(r[key]);
    if (r[key] === null && preliminary) return null;
    if (r[key] === null || r[key] === '') throw Error(`Missing number ${key}`);
    const n = Number(r[key]);
    if (!Number.isFinite(n)) throw Error(`Invalid number ${key}`);
    return n;
  }));
  const wb = Workbook.create();
  const sheet = wb.worksheets.add(item.type);
  sheet.showGridLines = false;
  const lastCol = col(spec.keys.length - 1), lastRow = rows.length + 3;
  const all = sheet.getRange(`A1:${lastCol}${lastRow + 1}`);
  all.format.font = { name: 'Arial', size: 11, color: '#202B38' };
  all.format.verticalAlignment = 'center';
  sheet.getRange('A1').values = [[item.type + (preliminary ? ' — ПРЕДВАРИТЕЛЬНЫЙ' : '')]];
  sheet.getRange('A1').format.font = { name: 'Arial', size: 16, bold: true, color: '#233A53' };
  sheet.getRange(`A1:${lastCol}1`).format.rowHeight = 29;
  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange('A2').values = [[`${item.base}, ${displayDate}` + (preliminary ? '. Все пробитые чеки включены. Жёлтые строки предварительные. Причины: лист «Проверка».' : '')]];
  sheet.getRange(`A2:${lastCol}2`).format.rowHeight = preliminary ? 36 : 25;
  sheet.getRange(`A2:${lastCol}2`).format.wrapText = true;
  if (preliminary) sheet.getRange(`A1:${lastCol}2`).format.fill = '#FFF2CC';
  sheet.getRange(`A3:${lastCol}3`).values = [spec.headers];
  if (rows.length) sheet.getRange(`A4:${lastCol}${lastRow}`).values = rows;
  if (item.type === 'Продажи') {
    // "Чеки" includes sale and refund receipts. Gross sales = net revenue +
    // refunds; average purchase excludes refund receipts from its denominator.
    rows.forEach((r, index) => {
      if (r[7] !== null) {
        const n = index + 4;
        sheet.getRange(`H${n}`).formulas = [[`=(C${n}+G${n})/(D${n}-F${n})`]];
      }
    });
  }
  spec.widths.forEach((w, i) => sheet.getRange(`${col(i)}1:${col(i)}${lastRow + 1}`).format.columnWidth = w);
  const head = sheet.getRange(`A3:${lastCol}3`);
  head.format = { fill: '#233A53', font: { name: 'Arial', size: 11, bold: true, color: '#FFFFFF' },
    wrapText: true, horizontalAlignment: 'center', verticalAlignment: 'center', rowHeight: 40 };
  head.format.borders = { insideVertical: { style: 'thin', color: '#FFFFFF' } };
  if (rows.length) {
    const body = sheet.getRange(`A4:${lastCol}${lastRow}`);
    body.format.wrapText = true;
    body.format.rowHeight = 34;
    body.format.horizontalAlignment = 'left';
    for (const i of spec.numeric) {
      sheet.getRange(`${col(i)}4:${col(i)}${lastRow + 1}`).format.horizontalAlignment = 'right';
      sheet.getRange(`${col(i)}4:${col(i)}${lastRow + 1}`).setNumberFormat(
        rows.every(r => r[i] === null || Number.isInteger(r[i])) ? '#,##0' : '#,##0.000');
    }
    if (!item.end_date && ['Продажи', 'Остатки'].includes(item.type)) sheet.getRange(`A4:A${lastRow}`).setNumberFormat('dd.mm.yyyy');
    if (item.type === 'Продажи') {
      for (const c of ['C','G','H']) sheet.getRange(`${c}4:${c}${lastRow+1}`).setNumberFormat('#,##0.00');
    }
    if (item.type === 'Бонусы') {
      sheet.getRange(`A4:A${lastRow}`).setNumberFormat('dd.mm.yyyy');
      sheet.getRange(`B4:B${lastRow}`).setNumberFormat('@');
      sheet.getRange(`C4:C${lastRow+1}`).setNumberFormat('#,##0.00');
    }
    body.format.autofitRows();
    for (let n = 4; n <= lastRow; n += 2) sheet.getRange(`A${n}:${lastCol}${n}`).format.fill = '#F2F5F8';
    const table = sheet.tables.add(`A3:${lastCol}${lastRow}`, true, spec.table);
    table.showFilterButton = true;
    if (item.type === 'Контроль') sheet.getRange(`E4:F${lastRow}`).setNumberFormat('dd.mm.yyyy hh:mm');
    if (preliminary) sheet.getRange(`A4:${lastCol}${lastRow}`).conditionalFormats.addCustom(`=$${lastCol}4="Предварительно"`, { fill: '#FFF2CC' });
  }
  const totalRow = lastRow + 1;
  const isAverage = i => item.type === 'Продажи' && i === 7;
  const purchaseCount = () => rows.reduce((sum,r) => sum + Number(r[3]) - Number(r[5]), 0);
  const incomplete = i => source.incomplete_scope || (isAverage(i)
    ? rows.some(r => [2,3,5,6].some(k => r[k] === null)) || purchaseCount() === 0
    : rows.some(r => r[i] === null));
  sheet.getRange(`A${totalRow}`).values = [[preliminary ? 'Итого по журналу (предварительно)' : 'Итого']];
  for (const i of spec.numeric) {
    const c = col(i);
    if (!incomplete(i)) sheet.getRange(`${c}${totalRow}`).formulas = [[isAverage(i)
      ? `=(C${totalRow}+G${totalRow})/(D${totalRow}-F${totalRow})`
      : rows.length ? `=SUM(${c}4:${c}${lastRow})` : '=0']];
  }
  sheet.getRange(`A${totalRow}:${lastCol}${totalRow}`).format = {
    font: { name: 'Arial', size: 11, bold: true }, rowHeight: 28,
    borders: { top: { style: 'thin', color: '#233A53' } },
  };
  if (rows.length > 12) sheet.freezePanes.freezeRows(3);
  if (preliminary) {
    const check = wb.worksheets.add('Проверка');
    check.showGridLines = false;
    check.getRange('A1').values = [['Почему отчёт предварительный']];
    check.getRange('A1').format.font = { name: 'Arial', size: 15, bold: true };
    check.getRange('A2:C2').merge();
    check.getRange('A2').values = [['Все пробитые чеки включены. Пустое число означает отсутствие подтверждённого показателя. Следующий запуск повторно сверит этот день.']];
    check.getRange('A2:C2').format.wrapText = true;
    check.getRange('A2:C2').format.rowHeight = 34;
    check.getRange('A3:C3').values = [['Магазин', 'Причина', 'Исходные значения']];
    const details = (source.issues || []).map(i => [i.shop, i.message, i.description || JSON.stringify(i.detail)]);
    if (details.length) check.getRange(`A4:C${details.length + 3}`).values = details.map(r => r.map(v => String(v).startsWith('=') ? "'" + v : v));
    [52,65,85].forEach((width, i) => check.getRange(`${col(i)}1:${col(i)}${details.length + 3}`).format.columnWidth = width);
    check.getRange('A3:C3').format = { fill: '#233A53', font: { name: 'Arial', size: 11, bold: true, color: '#FFFFFF' }, rowHeight: 28 };
    check.getRange(`A4:C${Math.max(4, details.length + 3)}`).format.wrapText = true;
    check.getRange(`A4:C${Math.max(4, details.length + 3)}`).format.autofitRows();
    check.freezePanes.freezeRows(3);
    sheet.getRange(`${lastCol}1`).values = [['См. лист «Проверка»']];
  }
  wb.recalculate();
  const totals = spec.numeric.map(i => sheet.getRange(`${col(i)}${totalRow}`).values[0][0]);
  for (const [j, i] of spec.numeric.entries()) {
    if (incomplete(i)) {
      if (totals[j] !== null && totals[j] !== '') throw Error('Incomplete total must remain blank');
      continue;
    }
    const expected = isAverage(i) ? rows.reduce((sum,r) => sum+r[2]+r[6], 0) / purchaseCount()
                                 : rows.reduce((sum, r) => sum + r[i], 0);
    if (Math.abs(Number(totals[j]) - expected) > 0.000001) throw Error('Workbook total differs from input');
  }
  const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 5 }, maxChars: 1500 });
  await fs.writeFile(path.join(outputDir, `inspection-${item.type}.txt`), errors.ndjson);
  if (preliminary) {
    const checkPreview = await wb.render({ sheetName: 'Проверка', range: 'A1:C7', scale: 1, format: 'png' });
    await fs.writeFile(path.join(outputDir, `preview-check-${item.type}.png`), new Uint8Array(await checkPreview.arrayBuffer()));
  }
  const preview = await wb.render({ sheetName: item.type, range: `A1:${lastCol}${Math.min(totalRow, 12)}`, scale: 1.5, format: 'png' });
  await fs.writeFile(path.join(root, `preview-${item.type}.png`), new Uint8Array(await preview.arrayBuffer()));
  if (manifest.preview_all) await fs.writeFile(path.join(outputDir, `preview-${item.base}-${item.type}.png`), new Uint8Array(await preview.arrayBuffer()));
  const filename = `${displayDate}_${item.base}_${item.type}.xlsx`;
  if (/[<>:"/\\|?*]/.test(filename)) throw Error('Unsafe filename');
  await (await SpreadsheetFile.exportXlsx(wb)).save(path.join(outputDir, filename));
  console.log(JSON.stringify({ filename, rows: rows.length, totals }));
}
