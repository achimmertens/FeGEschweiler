import fs from 'fs';
import path from 'path';
import playwright from 'playwright';
import { findLatestAusgabenJson } from './chart-common.js';

async function generateEinnahmenChartPlaywright(reportYear, options = {}) {
  const resultDir = path.join(process.cwd(), 'Daten', 'result');
  let year = reportYear;
  if (!year) {
    const best = findLatestAusgabenJson(resultDir);
    if (!best) throw new Error(`No Einnahmen_YYYY.json files found in ${resultDir}`);
    year = best.year;
  }

  const jsonPath = path.join(resultDir, `Einnahmen_${year}.json`);
  if (!fs.existsSync(jsonPath)) throw new Error(`JSON file not found: ${jsonPath}`);
  const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));

  const years = Object.keys(data).sort();
  const categoryTotals = {};
  years.forEach(y => { const row = data[y] || {}; Object.entries(row).forEach(([cat, val]) => { categoryTotals[cat] = (categoryTotals[cat] || 0) + Math.abs(val || 0); }); });
  const hasSonstiges = Object.prototype.hasOwnProperty.call(categoryTotals, 'Sonstiges');
  if (hasSonstiges) delete categoryTotals['Sonstiges'];
  const sortedCats = Object.entries(categoryTotals).sort((a,b)=>b[1]-a[1]).map(e=>e[0]);
  if (hasSonstiges) sortedCats.push('Sonstiges');

  const colors = options.colors || ['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899','#06B6D4','#84CC16','#F97316','#6366F1'];
  const datasetsForHtml = sortedCats.map((cat, idx) => ({ label: cat, data: years.map(y => Math.abs((data[y] && data[y][cat]) || 0)), backgroundColor: (cat === 'Sonstiges') ? '#A9A9A9' : colors[idx % colors.length], borderColor: '#ffffff', borderWidth: 1, stack: 'stack1' }));

  const dataPart = JSON.stringify({ labels: years, datasets: datasetsForHtml }, null, 2);
  const titlePart = JSON.stringify({ display: true, text: `Einnahmen ${year}`, font: { size: 30 } }, null, 2);
  const legendPart = JSON.stringify({ position: 'top' }, null, 2);
  const formatterFn = `function(value, context) { try { var label = (context && context.dataset && context.dataset.label) ? context.dataset.label : ''; var v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : 0); var n = Math.round(Number(v) || 0); return (label ? label + ': ' : '') + n.toLocaleString('de-DE') + ' €'; } catch (e) { return '' + value; } }`;
  const datalabelsPart = `{"color":"#ffffff","anchor":"center","align":"center","padding":{"left":6},"font":{"weight":"bold","size":12},"formatter":${formatterFn}}`;
  const optionsPart = `{"indexAxis":"x","plugins":{"title":${titlePart},"legend":${legendPart},"datalabels":${datalabelsPart}},"scales":{"x":{"stacked":true},"y":{"stacked":true}},"responsive":false,"maintainAspectRatio":true}`;
  const cfgSource = `const cfg = {"type":"bar","data":${dataPart},"options":${optionsPart}};`;
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0}canvas{display:block}</style></head><body><canvas id="chart" width="1200" height="800"></canvas><script src="https://cdn.jsdelivr.net/npm/chart.js"></script><script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script><script>${cfgSource}Chart.register(ChartDataLabels);const ctx=document.getElementById('chart').getContext('2d');const chart=new Chart(ctx,cfg);window.__chart_ready=true;</script></body></html>`;

  const outHtmlPath = path.join(resultDir, `Einnahmen_${year}.html`);
  fs.writeFileSync(outHtmlPath, html, 'utf8');

  const browser = await playwright.chromium.launch({ headless: !!options.headless, args: ['--start-maximized'] });
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
  await page.setContent(html, { waitUntil: 'networkidle' });
  await page.waitForFunction(() => window.__chart_ready === true).catch(()=>{});
  const canvas = await page.$('canvas#chart');
  const outPath = path.join(resultDir, `Einnahmen_${year}.png`);
  if (canvas) { const box = await canvas.boundingBox(); await page.screenshot({ path: outPath, clip: { x: Math.max(0, Math.floor(box.x)), y: Math.max(0, Math.floor(box.y)), width: Math.ceil(box.width), height: Math.ceil(box.height) } }); } else { await page.screenshot({ path: outPath, fullPage: true }); }
  await browser.close();
  return { html: outHtmlPath, png: outPath };
}

export { generateEinnahmenChartPlaywright };
