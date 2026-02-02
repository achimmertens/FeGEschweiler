import PptxGenJS from 'pptxgenjs';
import { COLORS, ensureResultDirectory, formatGermanNumber } from './utils.js';

function inferYearsFromRows(dataRows) {
  if (!dataRows || dataRows.length === 0) return [];
  const keys = Object.keys(dataRows[0]);
  return keys.filter(k => /^\d{4}$/.test(k)).sort();
}

function createIncomeSlide(pptx, sortedIncomeData) {
  const slide = pptx.addSlide();
  slide.addText('Einnahmen', {
    x: 0.5, y: 0.2, w: 9, h: 0.6, fontSize: 28, bold: true, color: '#' + COLORS.primary, align: 'center'
  });
  if (!sortedIncomeData) return;
  const years = inferYearsFromRows(sortedIncomeData.data || []);

  const chartData = (sortedIncomeData.data || []).map(row => ({
    name: row.category,
    labels: years,
    values: years.map(y => row[y] || 0)
  }));

  slide.addChart(pptx.ChartType.bar, chartData, {
    x: 0.5, y: 1.0, w: 9, h: 3.8,
    barGrouping: 'stacked', showLegend: true, legendPos: 'r', chartColors: COLORS.chartColors
  });

  // Table below
  const table = [ ['Kategorie', ...years] ];
  (sortedIncomeData.data || []).forEach(row => {
    const r = [ row.category, ...years.map(y => formatGermanNumber(row[y] || 0) + ' €') ];
    table.push(r);
  });
  // Gesamt
  const totalRow = ['Gesamt', ...years.map(y => formatGermanNumber((sortedIncomeData.totals && sortedIncomeData.totals[y]) || 0) + ' €')];
  table.push(totalRow);

  slide.addTable(table, { x: 0.5, y: 4.9, w: 9, h: 2.0, fontSize: 9, colW: [3.5, ...years.map(() => 1.0)], align: 'left' });
}

function createExpensesSlide(pptx, sortedExpensesData) {
  const slide = pptx.addSlide();
  slide.addText('Ausgaben', {
    x: 0.5, y: 0.2, w: 9, h: 0.6, fontSize: 28, bold: true, color: '#' + COLORS.primary, align: 'center'
  });
  if (!sortedExpensesData) return;
  const years = inferYearsFromRows(sortedExpensesData.data || []);

  const chartData = (sortedExpensesData.data || []).map(row => ({
    name: row.category,
    labels: years,
    values: years.map(y => row[y] || 0)
  }));

  slide.addChart(pptx.ChartType.bar, chartData, {
    x: 0.5, y: 1.0, w: 9, h: 3.8,
    barGrouping: 'stacked', showLegend: true, legendPos: 'r', chartColors: COLORS.chartColors
  });

  // Table below
  const table = [ ['Kategorie', ...years] ];
  (sortedExpensesData.data || []).forEach(row => {
    const r = [ row.category, ...years.map(y => formatGermanNumber(row[y] || 0) + ' €') ];
    table.push(r);
  });
  const totalRow = ['Gesamt', ...years.map(y => formatGermanNumber((sortedExpensesData.totals && sortedExpensesData.totals[y]) || 0) + ' €')];
  table.push(totalRow);
  slide.addTable(table, { x: 0.5, y: 4.9, w: 9, h: 2.0, fontSize: 9, colW: [3.5, ...years.map(() => 1.0)], align: 'left' });
}

function createExpensesPieSlide(pptx, sortedPieData) {
  const slide = pptx.addSlide();
  slide.addText('Ausgaben nach Kategorien', { x: 0.5, y: 0.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: '#' + COLORS.primary, align: 'center' });
  if (!sortedPieData || sortedPieData.length === 0) return;

  const pieData = [{ name: 'Ausgaben', labels: sortedPieData.map(p => p.name), values: sortedPieData.map(p => p.value) }];
  slide.addChart(pptx.ChartType.pie, pieData, { x: 0.8, y: 1.0, w: 4.5, h: 4.5, showLegend: true, legendPos: 'r', chartColors: COLORS.chartColors });

  // Details table
  const total = sortedPieData.reduce((s, p) => s + p.value, 0);
  const table = [ ['Kategorie', 'Betrag', 'Anteil'] ];
  sortedPieData.forEach(p => table.push([ p.name, formatGermanNumber(p.value) + ' €', ((p.value / total) * 100).toFixed(1) + ' %' ]));
  slide.addTable(table, { x: 5.6, y: 1.0, w: 3.4, h: 4.5, fontSize: 9, colW: [2.2, 1.0, 0.8], align: 'left' });
}

export function createPPT(sortedIncomeData, sortedExpensesData, sortedPieData) {
  const pptx = new PptxGenJS();
  // Title slide
  const title = pptx.addSlide();
  title.addText('Finanzübersicht', { x: 1, y: 1.2, fontSize: 36, color: '#' + COLORS.primary, bold: true });

  createIncomeSlide(pptx, sortedIncomeData);
  createExpensesSlide(pptx, sortedExpensesData);
  createExpensesPieSlide(pptx, sortedPieData);

  ensureResultDirectory();
  const output = `${process.cwd()}/Daten/result/Präsentation.pptx`;
  pptx.writeFile({ fileName: output }).then(() => console.log(`PPTX gespeichert: ${output}`));
}
