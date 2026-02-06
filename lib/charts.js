// lib/charts.js

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { readCSV, extractExpenses, ensureResultDirectory, readProfitLossReports, extractIncome, readBalanceReports, extractAccountBalances, COLORS, computeDeltaSchuldenForYears } from './utils.js';
import QuickChart from 'quickchart-js';

// __dirname replacement for ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Use shared CSV reader and extractors from utils (handles German number formats and semicolon delimiter)
// readCSV(filePath) -> returns array of rows
// extractExpenses(report) -> returns { category: amount } with absolute expense values

async function generateAusgabenJson(currentYear) {
    try {
        fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] generateAusgabenJson called with currentYear=${currentYear}\n`);
    } catch (e) {}
    const reportYear = currentYear - 1; // Year for the output file name
    const yearsToProcess = [currentYear - 3, currentYear - 2, currentYear - 1]; // Last 3 years

    let allExpensesByYear = {}; // { year: { category: amount, ... }, ... }

    // Prefer using readProfitLossReports which already parses all reports
    const reports = readProfitLossReports();
    for (const year of yearsToProcess) {
        try {
            const report = reports[String(year)] || [];
            try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] report rows for ${year}: ${report.length}\n`); } catch (e) {}
            const expenses = extractExpenses(report);
            try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Extracted ${Object.keys(expenses).length} expense categories for ${year}\n`); } catch (e) {}
            allExpensesByYear[year] = expenses;
        } catch (error) {
            console.error(`Error processing report for ${year}:`, error);
            allExpensesByYear[year] = {};
        }
    }

    console.log('All expenses by year:', JSON.stringify(allExpensesByYear, null, 2));
    const outputData = {};
    yearsToProcess.forEach(year => {
        outputData[year] = { ...(allExpensesByYear[year] || {}) };
    });
    const deltaSchulden = computeDeltaSchuldenForYears(yearsToProcess);
    yearsToProcess.forEach(year => {
        if (deltaSchulden[year] !== undefined) {
            outputData[year]['Schuldenabbau'] = Math.round((deltaSchulden[year] || 0) * 100) / 100;
        }
    });

    // Save to JSON file (ensure result dir exists)
    ensureResultDirectory();
    const outputFileName = `Ausgaben_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    fs.writeFileSync(outputPath, JSON.stringify(outputData, null, 2), 'utf8');
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] JSON data saved to ${outputPath}\n`); } catch (e) {}
}

// Alternative: generate JSON from the already computed sortedExpensesData returned by createExcelFile
function generateAusgabenJsonFromSorted(sortedExpensesData, years, currentYear) {
    const reportYear = currentYear - 1;
    // Build per-year map from sortedExpensesData.data rows
    const yearsToProcess = years.slice(-3).map(String);
    const allExpensesByYear = {};
    yearsToProcess.forEach(y => { allExpensesByYear[y] = {}; });

    (sortedExpensesData.data || []).forEach(row => {
        const category = row.category;
        yearsToProcess.forEach(y => {
            allExpensesByYear[y][category] = row[y] || 0;
        });
    });

    // Compute combined totals
    const combined = {};
    yearsToProcess.forEach(y => {
        Object.entries(allExpensesByYear[y] || {}).forEach(([cat, val]) => {
            combined[cat] = (combined[cat] || 0) + (val || 0);
        });
    });

    // Determine top 9 categories
    const sortedCats = Object.entries(combined).sort((a, b) => b[1] - a[1]).map(e => e[0]);
    const top9 = sortedCats.slice(0, 9);
    const rest = sortedCats.slice(9);

    const output = {};
    yearsToProcess.forEach(y => {
        const obj = {};
        top9.forEach(cat => { obj[cat] = allExpensesByYear[y][cat] || 0; });
        // Sonstiges sum
        const sonst = rest.reduce((s, c) => s + ((allExpensesByYear[y][c] || 0)), 0);
        obj['Sonstiges'] = sonst;
        // Special handling for Darlehensleistungen:
        // - For years >= 2024, set Darlehensleistungen to the (negative) value of 'Zinsen für Darlehen' from the reports
        // - For 2023, keep existing Darlehensleistungen but ensure it's negative
        try {
            const yearNum = parseInt(y, 10);
            const zinsen = allExpensesByYear[y]['Zinsen für Darlehen'] || 0;
            const existingDarlehen = allExpensesByYear[y]['Darlehensleistungen'] || 0;
            if (yearNum >= 2024) {
                obj['Darlehensleistungen'] = -Math.abs(zinsen || 0);
            } else {
                obj['Darlehensleistungen'] = -Math.abs(existingDarlehen || 0);
            }
        } catch (e) {
            obj['Darlehensleistungen'] = obj['Darlehensleistungen'] || 0;
        }
        // Remove 'Zinsen für Darlehen' from final output (we consolidate it into Darlehensleistungen)
        if (obj.hasOwnProperty('Zinsen für Darlehen')) delete obj['Zinsen für Darlehen'];
        output[y] = obj;
    });

    ensureResultDirectory();
    const outputFileName = `Ausgaben_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8');
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] JSON from sorted data saved to ${outputPath}\n`); } catch (e) {}
    // also write debug dump of output and meta info
    try {
        const dbgPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${reportYear}_debug.json`);
        fs.writeFileSync(dbgPath, JSON.stringify({ top9, restCount: rest.length, output }, null, 2), 'utf8');
        fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Debug JSON written to ${dbgPath}\n`);
    } catch (e) { try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] ERROR writing debug JSON: ${e.message}\n`); } catch (ee) {} }
    return outputPath;
}

// Export the function to be used by generate-presentation.js
export {
    generateAusgabenJson,
    generateAusgabenJsonFromSorted
};

// Export new income functions
export {
    generateEinnahmenJson,
    generateEinnahmenJsonFromSorted
};

// Generate stacked bar PNG from Ausgaben_YYYY.json
async function generateAusgabenChartFromJson(reportYear) {
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] generateAusgabenChartFromJson called for ${reportYear}\n`); } catch (e) {}
    const jsonPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${reportYear}.json`);
    if (!fs.existsSync(jsonPath)) {
        throw new Error(`JSON file not found: ${jsonPath}`);
    }
    const content = fs.readFileSync(jsonPath, 'utf8');
    const data = JSON.parse(content);

    // years are keys sorted
    const years = Object.keys(data).sort();

    // collect categories and totals
    const categoryTotals = {};
    years.forEach(y => {
        const row = data[y] || {};
        Object.entries(row).forEach(([cat, val]) => {
            categoryTotals[cat] = (categoryTotals[cat] || 0) + (Math.abs(val) || 0);
        });
    });

    // remove Sonstiges from sorting, will place on top
    const hasSonstiges = Object.prototype.hasOwnProperty.call(categoryTotals, 'Sonstiges');
    if (hasSonstiges) delete categoryTotals['Sonstiges'];

    // sort categories by total desc (largest first -> bottom)
    const sortedCats = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]).map(e => e[0]);
    // append Sonstiges last (top)
    if (hasSonstiges) sortedCats.push('Sonstiges');

    // build datasets
    const datasets = sortedCats.map((cat, idx) => {
        const values = years.map(y => {
            const v = (data[y] && data[y][cat]) || 0;
            // Darlehensleistungen stay negative (below zero); others positive
            const val = (cat === 'Darlehensleistungen') ? -Math.abs(v || 0) : Math.abs(v || 0);
            return val;
        });
        const color = (cat === 'Sonstiges') ? '#A9A9A9' : `#${COLORS.chartColors[idx % COLORS.chartColors.length]}`;
        const dsObj = {
            label: cat,
            data: values,
            backgroundColor: color,
            borderColor: '#ffffff',
            borderWidth: 1,
            stack: 'stack1'
        };
        // per-dataset datalabels formatter as string for QuickChart
        dsObj.datalabels = {
            color: '#ffffff',
            formatter: "function(value, context) { try { var label = (context && context.dataset && context.dataset.label) ? context.dataset.label : ''; var v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : 0); var formatted = Number(v).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return (label ? label + ': ' : '') + formatted; } catch (e) { return '' + value; } }",
            anchor: 'center',
            align: 'start',
            padding: { left: 6 },
            font: { weight: 'bold', size: 12 },
            display: "function(value, context) { try { var v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : 0); return Math.abs(v || 0) > 0.5; } catch (e) { return false; } }"
        };
        return dsObj;
    });

    // Chart.js config
    const latestYear = years[years.length - 1] || reportYear;
    const config = {
        type: 'bar',
        data: {
            labels: years,
            datasets: datasets
        },
        options: {
            indexAxis: 'x',
            plugins: {
                title: { display: true, text: `Ausgaben ${reportYear} nach Kategorien (Top ${Math.max(0, sortedCats.length - (hasSonstiges?1:0))} + Sonstiges)`, font: { size: 20 } },
                legend: { position: 'top' },
                datalabels: {
                    color: '#ffffff',
                    formatter: "function(value, context) { try { var label = ''; if (context && context.chart && context.chart.data && typeof context.datasetIndex === 'number') { var ds = context.chart.data.datasets[context.datasetIndex]; if (ds && ds.label) label = ds.label; } var v = (typeof value === 'number') ? value : (value && value.y !== undefined ? value.y : 0); var formatted = Number(v).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return (label ? label + ': ' : '') + formatted; } catch (e) { return '' + value; } }",
                    anchor: 'center',
                    align: 'start',
                    padding: { left: 6 },
                    clamp: true,
                    font: { weight: 'bold', size: 12 }
                }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    };

    // render with QuickChart (remote rendering)
    const width = 1200;
    const height = 800;
    const qc = new QuickChart();
    // QuickChart expects function bodies as strings; our config already contains the formatter as a string
    // Write debug dump of the config and datasets to debug.log to inspect what is sent to QuickChart
    try {
        const debugPath = path.join(process.cwd(), 'debug.log');
        function serialize(obj) {
            return JSON.stringify(obj, function(key, value) {
                if (typeof value === 'function') return value.toString();
                return value;
            }, 2);
        }
        fs.appendFileSync(debugPath, `[charts] QuickChart config for Ausgaben_${reportYear}:\n` + serialize(config) + '\n');
        // Also log datasets simplified
        const dsLog = datasets.map(ds => ({ label: ds.label, data: ds.data }));
        fs.appendFileSync(debugPath, `[charts] datasets preview:\n` + JSON.stringify(dsLog, null, 2) + '\n');
    } catch (e) { /* ignore logging errors */ }
    qc.setConfig(config);
    qc.setWidth(width);
    qc.setHeight(height);
    qc.setBackgroundColor('white');
    ensureResultDirectory();
    const outPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${reportYear}.png`);
    await qc.toFile(outPath);
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Ausgaben PNG written to ${outPath}\n`); } catch (e) {}
    return outPath;
}

export { generateAusgabenChartFromJson };

// --- Einnahmen JSON generator (analog zu Ausgaben) ---
function generateEinnahmenJsonFromSorted(sortedIncomeData, years, currentYear) {
    const reportYear = currentYear - 1;
    const yearsToProcess = years.slice(-3).map(String);
    const allIncomeByYear = {};
    yearsToProcess.forEach(y => { allIncomeByYear[y] = {}; });

    (sortedIncomeData.data || []).forEach(row => {
        const category = row.category;
        yearsToProcess.forEach(y => {
            allIncomeByYear[y][category] = row[y] || 0;
        });
    });

    // Compute combined totals
    const combined = {};
    yearsToProcess.forEach(y => {
        Object.entries(allIncomeByYear[y] || {}).forEach(([cat, val]) => {
            combined[cat] = (combined[cat] || 0) + (val || 0);
        });
    });

    // Determine top 9 categories
    const sortedCats = Object.entries(combined).sort((a, b) => b[1] - a[1]).map(e => e[0]);
    const top9 = sortedCats.slice(0, 9);
    const rest = sortedCats.slice(9);

    const output = {};
    yearsToProcess.forEach(y => {
        const obj = {};
        top9.forEach(cat => { obj[cat] = allIncomeByYear[y][cat] || 0; });
        const sonst = rest.reduce((s, c) => s + ((allIncomeByYear[y][c] || 0)), 0);
        if (rest.length > 0) obj['Sonstiges'] = sonst;
        output[y] = obj;
    });

    ensureResultDirectory();
    const outputFileName = `Einnahmen_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8');
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Einnahmen JSON from sorted data saved to ${outputPath}\n`); } catch (e) {}
    return outputPath;
}

async function generateEinnahmenJson(currentYear) {
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] generateEinnahmenJson called with currentYear=${currentYear}\n`); } catch (e) {}
    const reportYear = currentYear - 1;
    const yearsToProcess = [currentYear - 3, currentYear - 2, currentYear - 1];
    const reports = readProfitLossReports();
    const allIncomeByYear = {};
    for (const year of yearsToProcess) {
        try {
            const report = reports[String(year)] || [];
            try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] income report rows for ${year}: ${report.length}\n`); } catch (e) {}
            const incomes = extractIncome(report);
            try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Extracted ${Object.keys(incomes).length} income categories for ${year}\n`); } catch (e) {}
            allIncomeByYear[year] = incomes;
        } catch (error) {
            console.error(`Error processing income report for ${year}:`, error);
            allIncomeByYear[year] = {};
        }
    }

    // combine and top9
    const combined = {};
    for (const y in allIncomeByYear) {
        for (const cat in allIncomeByYear[y]) {
            combined[cat] = (combined[cat] || 0) + allIncomeByYear[y][cat];
        }
    }
    const sortedCats = Object.entries(combined).sort((a, b) => b[1] - a[1]).map(e => e[0]);
    const top9 = sortedCats.slice(0, 9);
    const rest = sortedCats.slice(9);

    const output = {};
    yearsToProcess.forEach(y => {
        const obj = {};
        top9.forEach(cat => { obj[cat] = allIncomeByYear[y] && allIncomeByYear[y][cat] ? allIncomeByYear[y][cat] : 0; });
        const sonst = rest.reduce((s, c) => s + ((allIncomeByYear[y] && allIncomeByYear[y][c]) || 0), 0);
        if (rest.length > 0) obj['Sonstiges'] = sonst;
        output[y] = obj;
    });

    ensureResultDirectory();
    const outputFileName = `Einnahmen_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    try { fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8'); fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Einnahmen JSON saved to ${outputPath}\n`); } catch (e) { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] ERROR writing Einnahmen JSON: ${e.message}\n`); }
    return outputPath;
}
