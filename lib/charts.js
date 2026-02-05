// lib/charts.js

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { readCSV, extractExpenses, ensureResultDirectory, readProfitLossReports, extractIncome, readBalanceReports, extractAccountBalances, parseGermanNumber, COLORS } from './utils.js';
import QuickChart from 'quickchart-js';

// __dirname replacement for ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Use shared CSV reader and extractors from utils (handles German number formats and semicolon delimiter)
// readCSV(filePath) -> returns array of rows
// extractExpenses(report) -> returns { category: amount } with absolute expense values

// Helper function to get top N categories and 'Sonstiges'
function getTopCategoriesAndOthers(allExpensesCombined, topN = 9) { // Changed to top 9 as requested
    const sortedCategories = Object.entries(allExpensesCombined)
        .sort(([, amountA], [, amountB]) => amountB - amountA);

    const topCategories = sortedCategories.slice(0, topN).map(([category, amount]) => ({ category, amount }));
    const otherAmount = sortedCategories.slice(topN).reduce((sum, [, amount]) => sum + amount, 0);

    console.log("Top Categories:", topCategories);
    console.log("Other Amount:", otherAmount);

    return { topCategories, otherAmount };
}

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

    // Combine expenses from all years to find top categories
    let combinedExpensesTotal = {};
    for (const year in allExpensesByYear) {
        for (const category in allExpensesByYear[year]) {
            combinedExpensesTotal[category] = (combinedExpensesTotal[category] || 0) + allExpensesByYear[year][category];
        }
    }

    console.log('All expenses by year:', JSON.stringify(allExpensesByYear, null, 2));
    console.log('Combined expenses total:', JSON.stringify(combinedExpensesTotal, null, 2));

    const { topCategories, otherAmount } = getTopCategoriesAndOthers(combinedExpensesTotal, 9); // Changed to top 9

    // Prepare data for JSON output
    const outputData = {};

    for (const year of yearsToProcess) {
        const yearExpenses = allExpensesByYear[year] || {};
        let currentYearOtherExpenses = 0;
        const yearOutput = {};

        // Populate data for top categories
        topCategories.forEach(catInfo => {
            yearOutput[catInfo.category] = yearExpenses[catInfo.category] || 0;
        });

        // Calculate 'Sonstiges' for the current year
        if (otherAmount > 0) {
            for (const category in yearExpenses) {
                if (!topCategories.some(cat => cat.category === category)) {
                    currentYearOtherExpenses += yearExpenses[category];
                }
            }
            yearOutput['Sonstiges'] = currentYearOtherExpenses;
        }
        // Remove 'Zinsen für Darlehen' from final output (we consolidate it into Darlehensleistungen)
        if (yearOutput.hasOwnProperty('Zinsen für Darlehen')) delete yearOutput['Zinsen für Darlehen'];
        outputData[year] = yearOutput;
    }

    // Save to JSON file (ensure result dir exists)
    ensureResultDirectory();
    const outputFileName = `Ausgaben_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    
    try {
        // insert Schuldenabbau (debt delta) computed from Bilanzberichte
        try {
            const balanceReports = readBalanceReports();
            const yrs = Object.keys(outputData).sort();
            let prevSum = null;
            for (const y of yrs) {
                const bal = extractAccountBalances(balanceReports[y] || []);
                // Output all authoritative income categories unmodified to ensure sums match CSV
                const output = {};
                yearsToProcess.forEach(y => {
                    const auth = authoritativeByYear[y] || {};
                    const obj = {};
                    Object.entries(auth).forEach(([cat, val]) => { obj[cat] = Number(val || 0); });
                    // If there are any categories not present (shouldn't), ensure Sonstiges at least 0
                    if (!obj['Sonstiges']) obj['Sonstiges'] = obj['Sonstiges'] || 0;
                    output[y] = obj;
                });
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
async function generateEinnahmenJson(currentYear) {
    try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] generateEinnahmenJson called with currentYear=${currentYear}\n`); } catch (e) {}
    const reportYear = currentYear - 1;
    const yearsToProcess = [currentYear - 3, currentYear - 2, currentYear - 1];
    const reports = readProfitLossReports();

    // Build authoritative income maps per year directly from parsed reports
    const authoritativeByYear = {};
    yearsToProcess.forEach(y => {
        const rep = reports[String(y)] || [];
        try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] income report rows for ${y}: ${rep.length}\n`); } catch(e){}
        authoritativeByYear[y] = extractIncome(rep);
    });

    // Determine top categories across all years
    const combined = {};
    yearsToProcess.forEach(y => {
        const map = authoritativeByYear[y] || {};
        Object.entries(map).forEach(([cat, val]) => { combined[cat] = (combined[cat] || 0) + Number(val || 0); });
    });
    const sortedCats = Object.entries(combined).sort((a,b)=>b[1]-a[1]).map(e=>e[0]);
    const top9 = sortedCats.slice(0,9);

    // Build output by taking authoritative categories and grouping rest into Sonstiges
    const output = {};
    yearsToProcess.forEach(y => {
        const auth = authoritativeByYear[y] || {};
        const obj = {};
        // include top9 explicitly
        top9.forEach(cat => { if (auth[cat] !== undefined) obj[cat] = Number(auth[cat]); });
        // compute Sonstiges as sum of remaining authoritative categories
        const sonst = Object.entries(auth).reduce((s, [cat, val]) => top9.includes(cat) ? s : s + Number(val || 0), 0);
        if (sonst > 0) obj['Sonstiges'] = Number(sonst.toFixed(2));
        // ensure at least zero
        if (!obj['Sonstiges']) obj['Sonstiges'] = obj['Sonstiges'] || 0;
        // include any categories not in top9 but you may want them visible? (we keep in Sonstiges)
        output[y] = obj;
    });

    ensureResultDirectory();
    const outputFileName = `Einnahmen_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    try { fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8'); fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Einnahmen JSON saved to ${outputPath}\n`); } catch (e) { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] ERROR writing Einnahmen JSON: ${e.message}\n`); }
    return outputPath;
}
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

    const output = {};

    for (const y of yearsToProcess) {
        const rep = reports[String(y)] || [];
        try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] income report rows for ${y}: ${rep.length}\n`); } catch (e) {}

        // Build authoritative map from CSV: include rows where Summe>0 or Kontoklasse indicates revenue
        const authMap = {};
        let authoritativeTotal = 0;
        for (const row of rep) {
            try {
                const raw = row.Summe || row['Summe'] || 0;
                const v = parseGermanNumber(raw);
                const kk = (row.Kontoklasse || row['Kontoklasse'] || '').toString().toLowerCase();
                const name = (row.Name || row['Name'] || '').toString();
                const isRevenueClass = kk.includes('betr. erträge') || kk.includes('umsatzerlöse') || kk.includes('umsatz');
                const isInterest = /haben-?zinsen|zinsen/i.test(name);
                if (v > 0 || isRevenueClass || isInterest) {
                    const key = (row.Name || row['Name'] || '').toString().trim() || '(unnamed)';
                    authMap[key] = (authMap[key] || 0) + Number(v || 0);
                    authoritativeTotal += Number(v || 0);
                }
            } catch (e) { /* ignore parse errors */ }
        }

        // normalize and round
        Object.keys(authMap).forEach(k => { authMap[k] = Number(Number(authMap[k] || 0).toFixed(2)); });
        const listedSum = Object.values(authMap).reduce((s,n) => s + Number(n || 0), 0);

        // Ensure Sonstiges fills any tiny rounding gap; otherwise set to 0
        const gap = Number((authoritativeTotal - listedSum).toFixed(2));
        if (Math.abs(gap) > 0.001) {
            authMap['Sonstiges'] = Number((authMap['Sonstiges'] || 0) + gap);
        } else if (!Object.prototype.hasOwnProperty.call(authMap, 'Sonstiges')) {
            authMap['Sonstiges'] = 0;
        }

        try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] year=${y} authoritativeTotal=${authoritativeTotal} listedSum=${listedSum} gap=${gap}\n`); } catch (e) {}
        output[y] = authMap;
    }

    ensureResultDirectory();
    const outputFileName = `Einnahmen_${reportYear}.json`;
    const outputPath = path.join(process.cwd(), 'Daten', 'result', outputFileName);
    try { fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8'); fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] Einnahmen JSON saved to ${outputPath}\n`); } catch (e) { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] ERROR writing Einnahmen JSON: ${e.message}\n`); }
    return outputPath;
}
