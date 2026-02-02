// lib/charts.js

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { readCSV, extractExpenses, ensureResultDirectory, readProfitLossReports, extractIncome } from './utils.js';

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
        fs.writeFileSync(outputPath, JSON.stringify(outputData, null, 2), 'utf8');
        fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] JSON data saved to ${outputPath}\n`);
    } catch (e) {
        fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[charts] ERROR writing JSON: ${e.message}\n`);
    }
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
