import fs from 'fs';
import path from 'path';

function findLatestAusgabenJson(dir) {
  if (!fs.existsSync(dir)) return null;
  const files = fs.readdirSync(dir);
  const re = /^Ausgaben_(\d{4})\.json$/;
  let best = null;
  for (const f of files) {
    const m = f.match(re);
    if (m) {
      const year = Number(m[1]);
      if (!best || year > best.year) best = { file: f, year };
    }
  }
  return best;
}

async function generateAusgabenChartPlaywright(reportYear, options = {}) {
  // determine file if reportYear not provided
  const resultDir = path.join(process.cwd(), 'Daten', 'result');
  let year = reportYear;
  if (!year) {
    const best = findLatestAusgabenJson(resultDir);
    if (!best) throw new Error(`No Ausgaben_YYYY.json files found in ${resultDir}`);
    year = best.year;
  }

  const jsonPath = path.join(resultDir, `Ausgaben_${year}.json`);
  if (!fs.existsSync(jsonPath)) {
    throw new Error(`JSON file not found: ${jsonPath}`);
  }
  console.log(`Reading data from ${jsonPath}`);
  const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));

  // Build chart.js config
  const years = Object.keys(data).sort();

  const categoryTotals = {};
  years.forEach(y => {
    const row = data[y] || {};
    Object.entries(row).forEach(([cat, val]) => {
      categoryTotals[cat] = (categoryTotals[cat] || 0) + Math.abs(val || 0);
    });
  });

  const hasSonstiges = Object.prototype.hasOwnProperty.call(categoryTotals, 'Sonstiges');
  if (hasSonstiges) delete categoryTotals['Sonstiges'];
  const sortedCats = Object.entries(categoryTotals).sort((a,b)=>b[1]-a[1]).map(e=>e[0]);
  if (hasSonstiges) sortedCats.push('Sonstiges');

  const datasets = sortedCats.map((cat, idx) => {
    const vals = years.map(y => {
      const v = (data[y] && data[y][cat]) || 0;
      return (cat === 'Darlehensleistungen') ? -Math.abs(v || 0) : Math.abs(v || 0);
    });
    const colors = options.colors || [
      '#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899','#06B6D4','#84CC16','#F97316','#6366F1'
    ];
    const color = (cat === 'Sonstiges') ? '#A9A9A9' : colors[idx % colors.length];
    return { label: cat, data: vals, backgroundColor: color, borderColor: '#ffffff', borderWidth:1, stack: 'stack1' };
  });

  const config = {
    type: 'bar',
    data: { labels: years, datasets },
    options: {
      indexAxis: 'x',
      plugins: {
title: { display: true, text: `Ausgaben ${year} nach Kategorien`, font: { size: 50 } },
          legend: { position: 'top' },
          datalabels: {
            color: '#ffffff',
            anchor: 'center',
            align: 'start',
            padding: { left: 6 },
            font: { weight: 'bold', size: 24 },
          // Formatter: show "Kategorie: Wert" with German number formatting
          formatter: function(value, context) {
            try {
              const label = (context && context.dataset && context.dataset.label) ? context.dataset.label : '';
              // Format number to integer and append Euro symbol
              const formattedValue = Math.round(Number(value) || 0).toLocaleString('de-DE');
              return (label ? label + ': ' : '') + formattedValue + ' €';
            } catch (e) { return '' + value; }
          },
          display: function(value, context) {
            try {
              const v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : 0);
              return Math.abs(v || 0) > 0.5;
            } catch (e) { return false; }
          }
        }
      },
      scales: { x: { stacked: true }, y: { stacked: true } },
      responsive: false,
maintainAspectRatio: true
    }
  };

  // Build a nicely formatted cfg source for embedding in HTML. We stringify
  // the data portion and insert a real JS formatter function for datalabels.
  const datasetsForHtml = datasets.map(ds => ({
    label: ds.label,
    data: (ds.data || []).map(v => Number(Number(v || 0).toFixed(2))),
    backgroundColor: ds.backgroundColor,
    borderColor: ds.borderColor,
    borderWidth: ds.borderWidth,
    stack: ds.stack
  }));

  const dataPart = JSON.stringify({ labels: years, datasets: datasetsForHtml }, null, 2);

  const titlePart = JSON.stringify({ display: true, text: `Ausgaben ${year} nach Kategorien`, font: { size: 30 } }, null, 2);
  const legendPart = JSON.stringify({ position: 'top' }, null, 2);

  const formatterFn = `function(value, context) {
              try {
                const label = context.dataset && context.dataset.label ? context.dataset.label : '';
                const v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : value);
              // Format number to integer and append Euro symbol
              const formattedValue = Math.round(Number(value) || 0).toLocaleString('de-DE');
              return (label ? label + ': ' : '') + formattedValue + ' €';
              } catch (e) { return '' + value; }
            }`;

  const datalabelsPart = `{
      "color": "#ffffff",
      "anchor": "center",
      "align": "center",
      "padding": { "left": 6 },
      "font": { "weight": "bold", "size": 16 },
      "formatter": ${formatterFn}
    }`;

  const optionsPart = `{
    "indexAxis": "x",
    "plugins": {
      "title": ${titlePart},
      "legend": ${legendPart},
      "datalabels": ${datalabelsPart}
    },
    "scales": { "x": { "stacked": true }, "y": { "stacked": true } },
    "responsive": false,
    "maintainAspectRatio": true
  }`;

  const cfgSource = `const cfg = {
      "type": "bar",
      "data": ${dataPart},
      "options": ${optionsPart}
    };`;

  // HTML template
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin: 0; }
    canvas { display: block; }
  </style>
</head>
<body>
  <canvas id="chart" width="1200" height="800"></canvas>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
  <script>
    ${cfgSource}
    // register plugin
    Chart.register(ChartDataLabels);
    const ctx = document.getElementById('chart').getContext('2d');
    const chart = new Chart(ctx, cfg);
    // signal ready
    window.__chart_ready = true;
  </script>
</body>
</html>`;
  // save HTML as an intermediate file so user can open it locally
  const outHtmlPath = path.join(resultDir, `Ausgaben_${year}.html`);
  fs.writeFileSync(outHtmlPath, html, 'utf8');
  console.log(`Wrote HTML preview to ${outHtmlPath}`);

  // dynamic import of playwright
  let playwright;
  try {
    playwright = await import('playwright');
  } catch (e) {
    throw new Error('Playwright is not installed. Please run: npm install -D playwright && npx playwright install');
  }

  // Determine headless mode: explicit option > env var > default true
  let headless;
  if (options.headless !== undefined) headless = !!options.headless;
  else if (process.env.PLAYWRIGHT_HEADLESS !== undefined) {
    const v = String(process.env.PLAYWRIGHT_HEADLESS).toLowerCase();
    headless = !(v === 'false' || v === '0' || v === 'no');
  } else {
    headless = true;
  }
  console.log(`Launching browser (headless=${headless})`);
  const browser = await playwright.chromium.launch({ headless, args: ['--start-maximized'] });
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });

  // capture console and page errors for better diagnostics
  const debugLog = path.join(process.cwd(), 'debug.log');
  const pageErrors = [];
  page.on('console', msg => {
    try { fs.appendFileSync(debugLog, `[chart-playwright][console] ${msg.type()}: ${msg.text()}\n`, 'utf8'); } catch(e) {}
  });
  page.on('pageerror', err => {
    pageErrors.push(err && err.message ? err.message : String(err));
    try { fs.appendFileSync(debugLog, `[chart-playwright][pageerror] ${err && err.stack ? err.stack : String(err)}\n`, 'utf8'); } catch(e) {}
  });

  // navigate to the saved file so it's inspectable in the browser if needed
  await page.goto('file://' + outHtmlPath, { waitUntil: 'networkidle' });
  // wait for chart to be ready
  await page.waitForFunction(() => window.__chart_ready === true, { timeout: 10000 }).catch(()=>{});

  if (pageErrors.length > 0) {
    // include page errors in thrown error for createPresentation to log
    await browser.close();
    throw new Error('Page errors: ' + pageErrors.join(' | '));
  }

  // screenshot canvas
  const canvas = await page.$('canvas#chart');
  let outPath;
  if (canvas) {
    const box = await canvas.boundingBox();
    if (box && box.width > 0 && box.height > 0) {
      outPath = path.join(resultDir, `Ausgaben_${year}.png`);
      await page.screenshot({ path: outPath, clip: { x: Math.max(0, Math.floor(box.x)), y: Math.max(0, Math.floor(box.y)), width: Math.ceil(box.width), height: Math.ceil(box.height) } });
    }
  }
  if (!outPath) {
    // fallback to full page screenshot
    outPath = path.join(resultDir, `Ausgaben_${year}.png`);
    await page.screenshot({ path: outPath, fullPage: true });
  }

  await browser.close();
  console.log(`Saved chart PNG to ${outPath}`);
  return { html: outHtmlPath, png: outPath };
}

// Function to generate income chart using Playwright
async function generateEinnahmenChartPlaywright(reportYear, options = {}) {
  const resultDir = path.join(process.cwd(), 'Daten', 'result');
  const jsonPath = path.join(resultDir, `Einnahmen_${reportYear}.json`);
  if (!fs.existsSync(jsonPath)) {
    throw new Error(`JSON file not found: ${jsonPath}`);
  }
  console.log(`Reading income data from ${jsonPath}`);
  const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));

  // Prepare data for the chart
  const years = Object.keys(data).sort();
  const chartData = [];

  // Assuming jsonData is an array of objects like: [{ category: '...', value: ... }]
  // You might need to adjust this based on your actual JSON structure
  // For income, we expect positive values.
  Object.entries(data).forEach(([year, categories]) => {
    Object.entries(categories).forEach(([category, value]) => {
      // Find existing entry or create new one
      let entry = chartData.find(item => item.name === category);
      if (!entry) {
        entry = { name: category, value: 0 };
        chartData.push(entry);
      }
      entry.value += Math.abs(value || 0); // Ensure positive value for income
    });
  });

  // Sort chart data by value descending for better visualization
  chartData.sort((a, b) => b.value - a.value);

  // Prepare data for Chart.js
  const chartJsData = chartData.map(item => ({
    name: item.name,
    value: item.value
  }));

  // Build datasets per category (one dataset per income category across years)
  const categoryTotals = {};
  years.forEach(y => {
    const row = data[y] || {};
    Object.entries(row).forEach(([cat, val]) => {
      categoryTotals[cat] = (categoryTotals[cat] || 0) + Math.abs(val || 0);
    });
  });

  const hasSonstiges = Object.prototype.hasOwnProperty.call(categoryTotals, 'Sonstiges');
  if (hasSonstiges) delete categoryTotals['Sonstiges'];
  const sortedCats = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]).map(e => e[0]);
  if (hasSonstiges) sortedCats.push('Sonstiges');

  const colors = options.colors || [
    '#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899','#06B6D4','#84CC16','#F97316','#6366F1'
  ];

  const datasetsForHtml = sortedCats.map((cat, idx) => {
    const vals = years.map(y => {
      const v = (data[y] && data[y][cat]) || 0;
      return Math.abs(v || 0);
    });
    const color = (cat === 'Sonstiges') ? '#A9A9A9' : colors[idx % colors.length];
    return {
      label: cat,
      data: vals.map(v => Number(Number(v || 0).toFixed(2))),
      backgroundColor: color,
      borderColor: '#ffffff',
      borderWidth: 1,
      stack: 'stack1'
    };
  });

  const dataPart = JSON.stringify({ labels: years, datasets: datasetsForHtml }, null, 2);
  const titlePart = JSON.stringify({ display: true, text: `Einnahmen ${reportYear}`, font: { size: 30 } }, null, 2);
  const legendPart = JSON.stringify({ position: 'top' }, null, 2);

  const formatterFn = `function(value, context) { try { var label = (context && context.dataset && context.dataset.label) ? context.dataset.label : ''; var v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : 0); var n = Math.round(Number(v) || 0); var parts = n.toString().split('.'); return (label ? label + ': ' : '') + n.toLocaleString('de-DE') + ' €'; } catch (e) { return '' + value; } }`;

  const datalabelsPart = `{
    "color": "#ffffff",
    "anchor": "center",
    "align": "center",
    "padding": { "left": 6 },
    "font": { "weight": "bold", "size": 18 },
    "formatter": ${formatterFn}
  }`;

  const optionsPart = `{
    "indexAxis": "x",
    "plugins": {
      "title": ${titlePart},
      "legend": ${legendPart},
      "datalabels": ${datalabelsPart}
    },
    "scales": { "x": { "stacked": true }, "y": { "stacked": true } },
    "responsive": false,
    "maintainAspectRatio": true
  }`;

  const cfgSource = `const cfg = { "type": "bar", "data": ${dataPart}, "options": ${optionsPart} };`;

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin: 0; }
    canvas { display: block; }
  </style>
</head>
<body>
  <canvas id="chart" width="1200" height="800"></canvas>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
  <script>
    ${cfgSource}
    Chart.register(ChartDataLabels);
    const ctx = document.getElementById('chart').getContext('2d');
    const chart = new Chart(ctx, cfg);
    window.__chart_ready = true;
  </script>
</body>
</html>`;

  // dynamic import of playwright
  let playwright;
  try {
    playwright = await import('playwright');
  } catch (e) {
    throw new Error('Playwright is not installed. Please run: npm install -D playwright && npx playwright install');
  }

  // Determine headless mode
  let headless;
  if (options.headless !== undefined) headless = !!options.headless;
  else if (process.env.PLAYWRIGHT_HEADLESS !== undefined) {
    const v = String(process.env.PLAYWRIGHT_HEADLESS).toLowerCase();
    headless = !(v === 'false' || v === '0' || v === 'no');
  } else {
    headless = true;
  }
  console.log(`Launching browser (headless=${headless})`);
  const browser = await playwright.chromium.launch({ headless, args: ['--start-maximized'] });
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });

  // capture console and page errors
  const debugLog = path.join(process.cwd(), 'debug.log');
  const pageErrors = [];
  page.on('console', msg => {
    try { fs.appendFileSync(debugLog, `[income-chart-playwright][console] ${msg.type()}: ${msg.text()}\n`, 'utf8'); } catch(e) {}
  });
  page.on('pageerror', err => {
    pageErrors.push(err && err.message ? err.message : String(err));
    try { fs.appendFileSync(debugLog, `[income-chart-playwright][pageerror] ${err && err.stack ? err.stack : String(err)}\n`, 'utf8'); } catch(e) {}
  });

  // save HTML preview for inspection
  const outHtmlPath = path.join(resultDir, `Einnahmen_${reportYear}.html`);
  try { fs.writeFileSync(outHtmlPath, html, 'utf8'); console.log(`Wrote HTML preview to ${outHtmlPath}`); } catch (e) {}
  await page.setContent(html, { waitUntil: 'networkidle' });
  await page.waitForFunction(() => window.__chart_ready === true, { timeout: 10000 }).catch(()=>{});

  if (pageErrors.length > 0) {
    await browser.close();
    throw new Error('Page errors: ' + pageErrors.join(' | '));
  }

  const canvas = await page.$('canvas#chart');
  let outPath;
  if (canvas) {
    const box = await canvas.boundingBox();
    if (box && box.width > 0 && box.height > 0) {
      outPath = path.join(resultDir, `Einnahmen_${reportYear}.png`);
      await page.screenshot({ path: outPath, clip: { x: Math.max(0, Math.floor(box.x)), y: Math.max(0, Math.floor(box.y)), width: Math.ceil(box.width), height: Math.ceil(box.height) } });
    }
  }
  if (!outPath) {
    outPath = path.join(resultDir, `Einnahmen_${reportYear}.png`);
    await page.screenshot({ path: outPath, fullPage: true });
  }

  await browser.close();
  console.log(`Saved income chart PNG to ${outPath}`);
  return { html: outHtmlPath, png: outPath };
}

export { generateAusgabenChartPlaywright, generateEinnahmenChartPlaywright };
