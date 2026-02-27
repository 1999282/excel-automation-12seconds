// Web Worker: Offloads heavy data processing from the main UI thread

self.onmessage = function (e) {
    const { rawData } = e.data;
    if (!rawData || !rawData.length) {
        self.postMessage({ type: 'error', message: "No data provided to worker." });
        return;
    }

    let report = {
        initialRows: rawData.length,
        duplicatesRemoved: 0,
        missingValuesFilled: 0,
        datesStandardized: 0,
        textStandardized: 0,
        currenciesFixed: 0,
        returnsFlagged: 0
    };

    self.postMessage({ type: 'progress', step: 2 }); // Detecting Columns...

    // 1. Detect Columns
    const sampleKeys = Object.keys(rawData[0] || {}).map(k => String(k).toLowerCase().trim());
    const cols = {
        id: findCol(sampleKeys, ['id', 'order', 'no']),
        date: findCol(sampleKeys, ['date', 'time', 'data']),
        customer: findCol(sampleKeys, ['customer', 'client', 'name', 'company', 'buyer']),
        category: findCol(sampleKeys, ['category', 'type', 'group', 'department']),
        product: findCol(sampleKeys, ['product', 'item', 'desc']),
        quantity: findCol(sampleKeys, ['qty', 'quantity', 'amount']),
        revenue: findCol(sampleKeys, ['revenue', 'sales', 'total', 'price', 'value']),
        region: findCol(sampleKeys, ['region', 'area', 'territory', 'location', 'country'])
    };

    if (!cols.revenue) {
        self.postMessage({ type: 'error', message: "Could not detect a numeric column (Revenue/Sales/Total). Tool requires this to analyze." });
        return;
    }

    let data = [...rawData];

    // 2. Remove Duplicates
    const uniqueSet = new Set();
    const uniqueData = [];
    data.forEach(row => {
        const hash = Object.values(row).join('|');
        if (!uniqueSet.has(hash)) {
            uniqueSet.add(hash);
            uniqueData.push(row);
        } else {
            report.duplicatesRemoved++;
        }
    });
    data = uniqueData;

    self.postMessage({ type: 'progress', step: 3 }); // Cleaning Data...

    // 3. Process in chunks to allow progress updates back to the UI
    const CHUNK_SIZE = 5000;
    let i = 0;

    function processChunk() {
        let max = Math.min(i + CHUNK_SIZE, data.length);
        for (; i < max; i++) {
            let row = data[i];
            let cleanRow = { ...row };

            // Clean Customer Names (German fixes)
            if (cols.customer && cleanRow[cols.customer]) {
                let name = String(cleanRow[cols.customer]).trim();
                let lower = name.toLowerCase();
                if (lower.includes('gmbh') || lower.includes('kg') || lower.includes('ag')) {
                    name = name.replace(/gmbh/i, 'GmbH')
                        .replace(/kg/i, 'KG')
                        .replace(/ag/i, 'AG');
                    report.textStandardized++;
                }
                if (name !== String(row[cols.customer])) report.textStandardized++;
                cleanRow[cols.customer] = name.charAt(0).toUpperCase() + name.slice(1);
            }

            // Fill Missing Categories
            if (cols.category) {
                if (!cleanRow[cols.category] || String(cleanRow[cols.category]).trim() === "") {
                    cleanRow[cols.category] = 'Uncategorized';
                    report.missingValuesFilled++;
                }
            }

            // Format Dates
            if (cols.date && cleanRow[cols.date]) {
                let dStr = String(cleanRow[cols.date]).trim();
                let parsedDate = parseMessyDate(dStr);
                if (parsedDate) {
                    cleanRow[cols.date] = parsedDate;
                    if (dStr !== parsedDate) report.datesStandardized++;
                }
            }

            // Clean Currencies/Numbers
            if (cols.revenue && cleanRow[cols.revenue]) {
                let val = String(cleanRow[cols.revenue]);
                let num = parseMessyNumber(val);
                if (!isNaN(num)) {
                    cleanRow[cols.revenue] = num;
                    if (String(num) !== val.trim()) report.currenciesFixed++;
                } else {
                    cleanRow[cols.revenue] = 0;
                }
            }

            if (cols.quantity && cleanRow[cols.quantity]) {
                let qty = parseInt(String(cleanRow[cols.quantity]).trim(), 10) || 0;
                cleanRow[cols.quantity] = qty;

                // Flag Returns
                if (qty < 0 || (cols.revenue && cleanRow[cols.revenue] < 0)) {
                    cleanRow['Is_Return'] = 'Yes';
                    cleanRow[cols.revenue] = Math.abs(cleanRow[cols.revenue] || 0) * -1; // Ensure negative revenue
                    report.returnsFlagged++;
                } else {
                    cleanRow['Is_Return'] = 'No';
                }
            }

            // Calculate Profit & Margin
            if (cols.revenue) {
                let rev = cleanRow[cols.revenue] || 0;
                let profit = rev * 0.45; // Simulated 45% margin
                cleanRow['Est_Profit'] = parseFloat(profit.toFixed(2));
                cleanRow['Profit_Margin_%'] = 45.0;
            }

            data[i] = cleanRow;
        }

        if (i < data.length) {
            let percent = Math.round((i / data.length) * 100);
            self.postMessage({ type: 'progressUpdate', percent: percent, step: 3 });
            setTimeout(processChunk, 1); // yield to event loop
        } else {
            finishProcessing();
        }
    }

    processChunk();

    function finishProcessing() {
        self.postMessage({ type: 'progress', step: 4 }); // Calculating Stats...

        let totalIssues = report.duplicatesRemoved + report.missingValuesFilled + report.datesStandardized + report.textStandardized + report.currenciesFixed;
        let score = Math.max(15, 100 - ((totalIssues / Math.max(1, report.initialRows)) * 30));
        if (score > 99) score = 99;
        if (totalIssues === 0) score = 100;
        report.qualityScore = Math.round(score);
        report.totalIssues = totalIssues;

        let stats = {
            totalRevenue: data.reduce((sum, r) => sum + (Number(r[cols.revenue]) || 0), 0),
            totalProfit: data.reduce((sum, r) => sum + (Number(r['Est_Profit']) || 0), 0),
            totalItems: data.reduce((sum, r) => sum + Math.max(0, Number(r[cols.quantity]) || 0), 0),
            returnsCount: report.returnsFlagged,
            finalRows: data.length
        };

        self.postMessage({ type: 'progress', step: 5 }); // Aggregating Data...

        let byRegion = aggregate(data, cols.region, cols.revenue);
        let byProduct = aggregate(data, cols.category || cols.product, cols.revenue);
        let byDate = aggregate(data, cols.date, cols.revenue);
        // Sort dates chronologically
        byDate = Object.fromEntries(Object.entries(byDate).sort());

        // Top 5 Customers
        let byCustomer = aggregate(data, cols.customer, cols.revenue);
        let topCustomersObj = Object.entries(byCustomer).sort((a, b) => b[1] - a[1]).slice(0, 5);
        let topCustomers = Object.fromEntries(topCustomersObj);

        // Profit Margins by Category
        let marginsByCat = {};
        if (cols.category) {
            let catRev = aggregate(data, cols.category, cols.revenue);
            let catProf = aggregate(data, cols.category, 'Est_Profit');
            for (let c in catRev) {
                marginsByCat[c] = catRev[c] ? (catProf[c] / catRev[c]) * 100 : 0;
            }
        }

        self.postMessage({
            type: 'complete',
            data: data,
            report: report,
            cols: cols,
            stats: stats,
            aggregations: { byRegion, byProduct, byDate, topCustomers, marginsByCat }
        });
    }
};

// --- Utility Functions (Duplicated from app.js for isolated worker thread) ---

function findCol(keys, keywords) {
    for (let kw of keywords) {
        let match = keys.find(k => k.includes(kw));
        if (match) {
            return Object.keys(rawData[0] || {}).find(orig => String(orig).toLowerCase().trim() === match);
        }
    }
    return null;
}

let rawData = []; // Hack to keep findCol working easily, we'll just set it
self.addEventListener('message', function (e) {
    if (e.data && e.data.rawData) {
        rawData = e.data.rawData;
    }
});

function parseMessyDate(str) {
    if (!str) return null;
    str = String(str).trim();
    if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;
    let parts;
    if (str.includes('.')) {
        parts = str.split('.');
        if (parts.length >= 3) {
            let p1 = parseInt(parts[0], 10);
            let p2 = parseInt(parts[1], 10);
            let p3 = parseInt(parts[2], 10);
            if (p3 < 100) p3 += 2000;
            if (p1 > 12) return `${p3}-${String(p2).padStart(2, '0')}-${String(p1).padStart(2, '0')}`;
            return `${p3}-${String(p1).padStart(2, '0')}-${String(p2).padStart(2, '0')}`;
        }
    } else if (str.includes('/')) {
        parts = str.split('/');
        if (parts.length >= 3) {
            let p1 = parseInt(parts[0], 10);
            let p2 = parseInt(parts[1], 10);
            let p3 = parseInt(parts[2], 10);
            if (p3 < 100) p3 += 2000;
            if (p1 > 12) return `${p3}-${String(p2).padStart(2, '0')}-${String(p1).padStart(2, '0')}`;
            return `${p3}-${String(p1).padStart(2, '0')}-${String(p2).padStart(2, '0')}`;
        }
    }
    let d = new Date(str);
    if (!isNaN(d)) return d.toISOString().split('T')[0];
    return str;
}

function parseMessyNumber(val) {
    if (val === null || val === undefined || val === '') return NaN;
    val = String(val).trim();
    val = val.replace(/[^0-9.,-]/g, '');
    let commaCount = (val.match(/,/g) || []).length;
    let dotCount = (val.match(/\./g) || []).length;
    let lastCommaIdx = val.lastIndexOf(',');
    let lastDotIdx = val.lastIndexOf('.');

    if (commaCount > 0 && dotCount > 0) {
        if (lastCommaIdx > lastDotIdx) val = val.replace(/\./g, '').replace(',', '.');
        else val = val.replace(/,/g, '');
    } else if (commaCount > 0 && dotCount === 0) {
        if (lastCommaIdx === val.length - 3 || lastCommaIdx === val.length - 2 || val.length - lastCommaIdx <= 3) val = val.replace(',', '.');
        else val = val.replace(/,/g, '');
    } else if (commaCount === 0 && dotCount > 1) {
        let parts = val.split('.');
        val = parts.slice(0, -1).join('') + '.' + parts[parts.length - 1];
    }
    return parseFloat(val);
}

function aggregate(data, groupCol, valCol) {
    if (!groupCol || !valCol) return {};
    return data.reduce((acc, row) => {
        let key = row[groupCol] || 'Unknown';
        let val = Number(row[valCol]) || 0;
        acc[key] = (acc[key] || 0) + val;
        return acc;
    }, {});
}
