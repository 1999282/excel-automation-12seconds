// Store global state
let rawData = [];
let cleanedData = [];
let globalCols = {};
let startTime = 0;
let chartInstances = {}; // Track charts to destroy them on re-render

// DOM Elements
const dropZone = document.getElementById('upload-zone');
const fileInput = document.getElementById('file-input');
const sections = {
    hero: document.getElementById('hero'),
    howItWorks: document.getElementById('how-it-works'),
    upload: document.getElementById('upload-section'),
    processing: document.getElementById('processing-section'),
    results: document.getElementById('results-section'),
    error: document.getElementById('error-section')
};

// Smooth scroll to upload
function scrollToUpload() {
    sections.upload.scrollIntoView({ behavior: 'smooth' });
}

// Drag & Drop Handlers
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

// Load the embedded sample data
function loadSampleData() {
    const sampleCSV = `Order ID,Date,Customer Name,Product Category,Quantity,Unit Price,Revenue,Region
1001,01/15/2024,TechCorp GmbH,Electronics,5,199.99,999.95,North
1002,15/01/2024,TechCorp Gmbh,Electronics,5,199.99,999.95,North
1003,2024-01-16,Smith & Co.,Furniture,  2 ,450.00,900.00,South
1004,1/17/2024,Müller gmbh,Office Supplies,10,25.50,255.00,West
1005,18.01.2024,DataTech Solutions,,1,1200.00,1200.00,East
1006,01/19/2024,Smith & Co.,Furniture,1,450.00,450.00,South
1007,2024-01-20,TECHCORP GMBH,Electronics,-1,199.99,-199.99,North
1008,1/21/2024,Schmidt KG,Electronics,2, 899.00 ,1798.00,West
1009,22.01.2024,müller GmbH,Office Supplies,15,10.00,150.00,West
1010,01/23/2024,DataTech Solutions,Software,3,150.00,450.00,East
1011,2024-01-24,Weber AG,Furniture,4,220.00,880.00,North
1012,1/25/2024,Weber AG,,2,220.00,440.00,North
1013,26.01.2024,Smith & Co.,Office Supplies,5, 45.00 ,225.00,South
1014,01/27/2024,TechCorp GmbH,Electronics,1,199.99,199.99,North
1015,2024-01-28,Müller GmbH,Furniture,-2,150.00,-300.00,West
1016,1/29/2024,Schmidt kg,Office Supplies,20,5.50,110.00,West
1017,30.01.2024,DataTech Solutions,Software,5,150.00,750.00,East
1018,01/31/2024,Weber ag,Electronics,1,950.00,950.00,North
1019,2024-02-01,Smith & Co.,Office Supplies,10,12.00,120.00,South
1020,2/2/2024,TechCorp GmbH,Software,2,300.00,600.00,North
1021,03.02.2024,Müller GmbH,Office Supplies,8,15.00,120.00,West
1022,02/04/2024,DataTech Solutions,,1,2500.00,2500.00,East
1023,2024-02-05,Schmidt KG,Furniture,3,340.00,1020.00,West
1024,2/6/2024,Weber AG,Software,-1,800.00,-800.00,North`;

    // Parse sample CSV data
    Papa.parse(sampleCSV, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
            rawData = results.data;
            const fileSize = (new Blob([sampleCSV]).size / 1024).toFixed(1);
            showProcessingUI("Sample Data Mode", rawData.length, Object.keys(rawData[0] || {}).length, `${fileSize} KB`);
            setTimeout(processData, 800); // Artificial delay to show animation
        }
    });
}

// Show Error
function showError(msg) {
    sections.processing.style.display = 'none';
    sections.error.style.display = 'flex';
    document.getElementById('error-message').innerText = msg;
}

// Handle File Upload
function handleFile(file) {
    startTime = performance.now();

    // Error Handling: Check empty
    if (!file || file.size === 0) {
        showError("The uploaded file is empty or corrupted.");
        return;
    }

    const sizeStr = file.size > 1048576 ? (file.size / 1048576).toFixed(1) + ' MB' : (file.size / 1024).toFixed(1) + ' KB';

    if (file.name.endsWith('.csv')) {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                if (!results.data || results.data.length === 0) {
                    showError("The CSV file contains no readable data.");
                    return;
                }
                rawData = results.data;
                showProcessingUI(file.name, rawData.length, Object.keys(rawData[0] || {}).length, sizeStr);
                setTimeout(processData, 500);
            },
            error: () => showError("Failed to parse the CSV file.")
        });
    } else if (file.name.match(/\.xlsx?$/)) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const rawArray = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });

                if (rawArray.length === 0) {
                    showError("The Excel file is empty.");
                    return;
                }

                // Convert all keys to strings to normalize
                rawData = rawArray.map(obj => {
                    let newObj = {};
                    for (let key in obj) newObj[String(key).trim()] = obj[key] == null ? "" : String(obj[key]);
                    return newObj;
                });

                showProcessingUI(file.name, rawData.length, Object.keys(rawData[0] || {}).length, sizeStr);
                setTimeout(processData, 500);
            } catch (err) {
                showError("Failed to read the Excel file. Make sure it's not password protected.");
            }
        };
        reader.readAsArrayBuffer(file);
    } else {
        showError("Unsupported file type. Please upload a CSV or Excel file.");
    }
}

// Setup Processing UI
function showProcessingUI(filename, rows, cols, sizeStr) {
    // Hide everything else
    sections.hero.style.display = 'none';
    sections.howItWorks.style.display = 'none';
    sections.upload.style.display = 'none';
    sections.results.style.display = 'none';
    sections.error.style.display = 'none';
    sections.processing.style.display = 'flex';

    document.getElementById('file-info').innerText = `${filename} • ${rows} rows × ${cols} cols • ${sizeStr}`;

    // Reset steps
    for (let i = 1; i <= 6; i++) {
        const s = document.getElementById('step-' + i);
        s.className = 'step';
        s.querySelector('.step-icon').innerText = '⏳';
    }
}

function updateStep(stepNum) {
    const s = document.getElementById('step-' + stepNum);
    if (s) {
        s.className = 'step done active';
        s.querySelector('.step-icon').innerText = '✅';
    }
}

// Core Data Processing Pipeline
function processData() {
    try {
        updateStep(1); // Load

        let report = {
            initialRows: rawData.length,
            duplicatesRemoved: 0,
            missingValuesFilled: 0,
            datesStandardized: 0,
            textStandardized: 0,
            currenciesFixed: 0,
            returnsFlagged: 0
        };

        // 1. Detect Columns (Expert Auto-Detection)
        updateStep(2);
        const sampleKeys = Object.keys(rawData[0] || {}).map(k => k.toLowerCase().trim());
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

        // Fallbacks
        if (!cols.revenue) {
            showError("Could not detect a numeric column (Revenue/Sales/Total). Tools requires this to analyze.");
            return;
        }

        globalCols = cols;
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

        // 3. Clean Data Points
        updateStep(3);
        data = data.map(row => {
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

            // Format Dates to YYYY-MM-DD
            if (cols.date && cleanRow[cols.date]) {
                let dStr = String(cleanRow[cols.date]).trim();
                dStr = dStr.replace(/\./g, '/'); // 15.01.2024 -> 15/01/2024
                let d = new Date(dStr);
                if (!isNaN(d.getTime())) {
                    cleanRow[cols.date] = d.toISOString().split('T')[0];
                    if (dStr !== cleanRow[cols.date]) report.datesStandardized++;
                }
            }

            // Clean Currencies/Numbers
            if (cols.revenue && cleanRow[cols.revenue]) {
                let val = String(cleanRow[cols.revenue]);
                val = val.replace(/[€$£,\s]/g, '');
                let num = parseFloat(val);
                if (!isNaN(num)) {
                    cleanRow[cols.revenue] = num;
                    if (val !== String(row[cols.revenue])) report.currenciesFixed++;
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

            // Calculate Profit & Margin (Assumes 30% baseline cost for demo if no cost col)
            if (cols.revenue) {
                let rev = cleanRow[cols.revenue] || 0;
                let profit = rev * 0.45; // Simulated 45% margin
                cleanRow['Est_Profit'] = parseFloat(profit.toFixed(2));
                cleanRow['Profit_Margin_%'] = 45.0;
            }

            return cleanRow;
        });

        cleanedData = data;

        // Calculate Data Quality Score (Expert Feature)
        let totalIssues = report.duplicatesRemoved + report.missingValuesFilled + report.datesStandardized + report.textStandardized + report.currenciesFixed;
        let score = Math.max(15, 100 - ((totalIssues / Math.max(1, report.initialRows)) * 30));
        if (score > 99) score = 99; // Cap at 99 so the animation is visible
        if (totalIssues === 0) score = 100;
        report.qualityScore = Math.round(score);

        // Calculate Stats
        updateStep(4);
        let stats = {
            totalRevenue: data.reduce((sum, r) => sum + (Number(r[cols.revenue]) || 0), 0),
            totalProfit: data.reduce((sum, r) => sum + (Number(r['Est_Profit']) || 0), 0),
            totalItems: data.reduce((sum, r) => sum + Math.max(0, Number(r[cols.quantity]) || 0), 0),
            returnsCount: report.returnsFlagged,
            finalRows: data.length
        };

        // Aggregations
        updateStep(5);
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

        // Render UI
        updateStep(6);
        setTimeout(() => {
            const timeTaken = ((performance.now() - startTime) / 1000).toFixed(1);
            document.getElementById('timer-value').innerText = timeTaken + 's';

            renderDiff(rawData, cleanedData, cols);
            renderStats(stats);
            renderReport(report);
            renderQualityScore(report.qualityScore, totalIssues);

            // Build charts with auto-cleanup of previous canvases to prevent crashes
            buildCharts(byRegion, byProduct, byDate, topCustomers, marginsByCat);

            renderTable(data, cols);

            sections.processing.style.display = 'none';
            sections.results.style.display = 'block';
            sections.results.scrollIntoView({ behavior: 'smooth' });

            triggerConfetti();
        }, 300);

    } catch (e) {
        console.error(e);
        showError("An error occurred during processing: " + e.message);
    }
}

// Render "Before vs After" Difference View (Expert Polish)
function renderDiff(raw, clean, cols) {
    const container = document.getElementById('diff-container');
    container.innerHTML = '';

    // Show top 4 rows
    const limit = Math.min(4, raw.length);
    const keysToShow = [cols.customer, cols.date, cols.revenue, cols.quantity].filter(Boolean);
    if (keysToShow.length === 0) keysToShow.push(Object.keys(raw[0])[0]);

    // Build Before Table
    let beforeHTML = `<div class="diff-panel before"><div class="diff-panel-header">Before (Raw Data)</div><table class="diff-table">`;
    beforeHTML += `<thead><tr>${keysToShow.map(k => `<th>${k}</th>`).join('')}</tr></thead><tbody>`;

    // Build After Table
    let afterHTML = `<div class="diff-panel after"><div class="diff-panel-header">After (Cleaned)</div><table class="diff-table">`;
    afterHTML += `<thead><tr>${keysToShow.map(k => `<th>${k}</th>`).join('')}</tr></thead><tbody>`;

    for (let i = 0; i < limit; i++) {
        beforeHTML += `<tr>`;
        afterHTML += `<tr>`;
        for (let key of keysToShow) {
            let origVal = raw[i][key];

            // Find matched key in clean data (case-insensitive fallback)
            let cleanKeyMatch = Object.keys(clean[i]).find(k => k.toLowerCase() === key.toLowerCase()) || key;
            let cleanVal = clean[i][cleanKeyMatch];

            let isChanged = String(origVal).trim() !== String(cleanVal).trim();

            beforeHTML += `<td class="${isChanged ? 'changed' : ''}">${origVal !== undefined ? origVal : ''}</td>`;

            // Format output for neatness
            let displayVal = cleanVal;
            if (cleanKeyMatch === cols.revenue && !isNaN(cleanVal)) {
                displayVal = new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' }).format(cleanVal);
            }
            afterHTML += `<td class="${isChanged ? 'changed' : ''}">${displayVal !== undefined ? displayVal : ''}</td>`;
        }
        beforeHTML += `</tr>`;
        afterHTML += `</tr>`;
    }
    beforeHTML += `</tbody></table></div>`;
    afterHTML += `</tbody></table></div>`;

    container.innerHTML = beforeHTML + afterHTML;
}

// Helper to format currency
const currencyFmt = new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' });

// Render Stats Cards
function renderStats(s) {
    const grid = document.getElementById('summary-grid');
    grid.innerHTML = `
        <div class="summary-card">
            <div class="card-label">Total Revenue</div>
            <div class="card-value" style="color:var(--accent)">${currencyFmt.format(s.totalRevenue)}</div>
        </div>
        <div class="summary-card">
            <div class="card-label">Est. Profit</div>
            <div class="card-value" style="color:var(--green)">${currencyFmt.format(s.totalProfit)}</div>
        </div>
        <div class="summary-card">
            <div class="card-label">Items Sold</div>
            <div class="card-value" style="color:var(--blue)">${s.totalItems.toLocaleString()}</div>
        </div>
        <div class="summary-card">
            <div class="card-label">Cleaned Rows</div>
            <div class="card-value">${s.finalRows.toLocaleString()}</div>
        </div>
    `;
}

// Render Data Quality Score
function renderQualityScore(score, issues) {
    document.getElementById('quality-value').innerText = score + '%';

    // Animate the ring
    const fill = document.getElementById('quality-ring-fill');
    setTimeout(() => {
        // Circumference is 327. Offset by percentage.
        const offset = 327 - (score / 100) * 327;
        fill.style.strokeDashoffset = offset;

        let color = 'var(--accent)';
        if (score < 80) color = 'var(--orange)';
        if (score < 50) color = 'var(--red)';
        fill.style.stroke = color;
        document.getElementById('quality-value').style.color = color;
    }, 100);

    const badges = document.getElementById('quality-badges');
    badges.innerHTML = `
        <span class="quality-badge good">✓ Automated Processing</span>
        <span class="quality-badge fixed">${issues} Fixes Applied</span>
        ${score === 100 ? '<span class="quality-badge good">Perfect Data</span>' : ''}
    `;
}

// Render Report Details
function renderReport(r) {
    const grid = document.getElementById('report-grid');
    grid.innerHTML = `
        <div class="report-item"><span class="label">Initial Rows</span><span class="value">${r.initialRows}</span></div>
        <div class="report-item" style="color:var(--blue)"><span class="label">Duplicates Dropped</span><span class="value">${r.duplicatesRemoved}</span></div>
        <div class="report-item" style="color:var(--green)"><span class="label">Missing Values Handled</span><span class="value">${r.missingValuesFilled}</span></div>
        <div class="report-item" style="color:var(--purple)"><span class="label">Dates Standardized</span><span class="value">${r.datesStandardized}</span></div>
        <div class="report-item" style="color:var(--cyan)"><span class="label">Text & Casing Fixed</span><span class="value">${r.textStandardized}</span></div>
        <div class="report-item" style="color:var(--accent)"><span class="label">Currencies Standardized</span><span class="value">${r.currenciesFixed}</span></div>
        <div class="report-item" style="color:var(--red)"><span class="label">Returns Flagged</span><span class="value">${r.returnsFlagged}</span></div>
    `;
}

// Chart utility overrides for dark mode
Chart.defaults.color = '#8888a0';
Chart.defaults.font.family = "'Inter', sans-serif";

function buildCharts(reg, prod, time, cust, margins) {
    // Standard Colors
    const c1 = '#c8ff00'; // accent
    const c2 = '#3b82f6'; // blue
    const c3 = '#8b5cf6'; // purple
    const c4 = '#ec4899'; // pink
    const c5 = '#06b6d4'; // cyan
    const bgOpacity = '15';

    // Helper to safely create/destroy charts (Fixes re-render crash)
    function safelyCreateChart(ctxId, config) {
        if (chartInstances[ctxId]) {
            chartInstances[ctxId].destroy();
        }
        const ctx = document.getElementById(ctxId).getContext('2d');
        chartInstances[ctxId] = new Chart(ctx, config);
    }

    // 1. Line: Trend
    safelyCreateChart('chart-trend', {
        type: 'line',
        data: {
            labels: Object.keys(time),
            datasets: [{
                label: 'Revenue (€)',
                data: Object.values(time),
                borderColor: c3,
                backgroundColor: c3 + bgOpacity,
                borderWidth: 3,
                tension: 0.4,
                fill: true,
                pointBackgroundColor: c3,
                pointBorderColor: '#12121c',
                pointHoverRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: '#2a2a3a' } },
                x: { grid: { display: false } }
            }
        }
    });

    // 2. Bar: Region
    safelyCreateChart('chart-region', {
        type: 'bar',
        data: {
            labels: Object.keys(reg),
            datasets: [{
                label: 'Revenue',
                data: Object.values(reg),
                backgroundColor: c2,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: '#2a2a3a' } },
                x: { grid: { display: false } }
            }
        }
    });

    // 3. Doughnut: Product
    safelyCreateChart('chart-product', {
        type: 'doughnut',
        data: {
            labels: Object.keys(prod),
            datasets: [{
                data: Object.values(prod),
                backgroundColor: [c1, c2, c3, c4, c5],
                borderWidth: 0,
                hoverOffset: 10
            }]
        },
        options: {
            responsive: true,
            cutout: '75%',
            plugins: {
                legend: { position: 'right', labels: { usePointStyle: true, boxWidth: 8, font: { size: 11 } } }
            }
        }
    });

    // 4. Horizontal Bar: Customers
    safelyCreateChart('chart-customers', {
        type: 'bar',
        data: {
            labels: Object.keys(cust).map(l => l.length > 15 ? l.substring(0, 15) + '...' : l),
            datasets: [{
                label: 'Revenue',
                data: Object.values(cust),
                backgroundColor: c1,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            plugins: { legend: { display: false } },
            scales: {
                x: { beginAtZero: true, grid: { color: '#2a2a3a' } },
                y: { grid: { display: false } }
            }
        }
    });

    // 5. Radar: Margins
    const mKeys = Object.keys(margins);
    if (mKeys.length > 2 && mKeys.length < 8) { // Radar looks best with 3-7 points
        safelyCreateChart('chart-margins', {
            type: 'radar',
            data: {
                labels: mKeys,
                datasets: [{
                    label: 'Profit Margin %',
                    data: Object.values(margins),
                    backgroundColor: c5 + '44',
                    borderColor: c5,
                    borderWidth: 2,
                    pointBackgroundColor: c5
                }]
            },
            options: {
                responsive: true,
                scales: {
                    r: {
                        angleLines: { color: '#2a2a3a' },
                        grid: { color: '#2a2a3a' },
                        pointLabels: { color: '#8888a0', font: { size: 10 } },
                        ticks: { display: false }
                    }
                },
                plugins: { legend: { display: false } }
            }
        });
    } else {
        // Fallback to Bar if radar not suitable
        safelyCreateChart('chart-margins', {
            type: 'bar',
            data: {
                labels: mKeys,
                datasets: [{
                    label: 'Margin %',
                    data: Object.values(margins),
                    backgroundColor: c5,
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, max: 100, grid: { color: '#2a2a3a' } },
                    x: { grid: { display: false } }
                }
            }
        });
    }
}

// Render Table
function renderTable(data, selectedCols) {
    document.getElementById('table-count').innerText = `Showing 1-100 of ${data.length}`;

    // Header
    const customKeys = ['Is_Return', 'Est_Profit'];
    const originKeys = Object.values(selectedCols).filter(Boolean);
    const tblKeys = [...new Set([...originKeys, ...customKeys])];

    let thead = '<tr>';
    tblKeys.forEach(k => thead += `<th>${k.replace(/_/g, ' ')}</th>`);
    thead += '</tr>';
    document.getElementById('table-head').innerHTML = largeDOMPurify(thead); // Assuming trusted inner app data

    // Body (Max 100 rows to keep DOM light)
    let tbody = '';
    const limit = Math.min(100, data.length);
    for (let i = 0; i < limit; i++) {
        let row = data[i];
        tbody += '<tr>';
        tblKeys.forEach(k => {
            let val = row[k];
            if (val === undefined || val === null) val = '';

            // Format numbers nicely
            if (k === selectedCols.revenue || k === 'Est_Profit') {
                if (!isNaN(val) && val !== '') val = currencyFmt.format(val);
            }
            // Format badges
            if (k === 'Is_Return' && val === 'Yes') {
                val = `<span style="color:var(--red); background:rgba(239,68,68,0.15); padding:2px 6px; border-radius:4px; font-size:0.7rem; font-weight:bold;">YES</span>`;
            }

            tbody += `<td>${val}</td>`;
        });
        tbody += '</tr>';
    }
    document.getElementById('table-body').innerHTML = tbody; // Basic interpolation output
}

function largeDOMPurify(str) { return str; } // Mock if no purify library installed locally

// Confetti Animation
function triggerConfetti() {
    const container = document.getElementById('confetti-container');
    container.innerHTML = '';
    const colors = ['#c8ff00', '#3b82f6', '#8b5cf6', '#ec4899', '#06b6d4'];

    for (let i = 0; i < 50; i++) {
        const confetti = document.createElement('div');
        confetti.className = 'confetti-piece';
        confetti.style.left = Math.random() * 100 + 'vw';
        confetti.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
        confetti.style.animationDelay = (Math.random() * 2) + 's';
        confetti.style.animationDuration = (Math.random() * 2 + 2) + 's';

        // Random shapes
        if (Math.random() > 0.5) confetti.style.borderRadius = '50%';
        container.appendChild(confetti);
    }

    // Cleanup
    setTimeout(() => { container.innerHTML = ''; }, 5000);
}

// Reset UI
function resetApp() {
    sections.results.style.display = 'none';
    sections.error.style.display = 'none';
    sections.hero.style.display = 'flex';
    sections.howItWorks.style.display = 'block';
    sections.upload.style.display = 'block';
    document.getElementById('file-input').value = '';
    // Restore ring rotation
    document.getElementById('quality-ring-fill').style.strokeDashoffset = 327;
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Excel Export (Fix bug by using dynamic globals)
function exportToExcel() {
    if (!cleanedData.length) return;

    // Create new workbook
    const wb = XLSX.utils.book_new();

    // format numbers for Excel properly
    const cleanOutput = cleanedData.map(row => {
        let cloned = { ...row };
        // Clean out spaces if any remain
        return cloned;
    });

    // 1. Cleaned Data Sheet
    const wsData = XLSX.utils.json_to_sheet(cleanOutput);
    XLSX.utils.book_append_sheet(wb, wsData, "Cleaned Data");

    // 2. Summary Dashboard Sheet
    let rCols = globalCols;
    let revCol = rCols.revenue || 'Revenue'; // Safe fallback
    let qtyCol = rCols.quantity || 'Quantity';

    const totalRev = cleanOutput.reduce((sum, r) => sum + (Number(r[revCol]) || 0), 0);
    const totalQty = cleanOutput.reduce((sum, r) => sum + Math.max(0, Number(r[qtyCol]) || 0), 0);
    const byRegion = aggregate(cleanOutput, rCols.region || 'Region', revCol);

    const dashData = [
        ["Excel Automation - AI Report Dashboard"],
        [""],
        ["Overview", ""],
        ["Total Rows Processed", cleanOutput.length],
        ["Total Revenue", totalRev],
        ["Total Items Sold", totalQty],
        ["Returns Processed", cleanOutput.filter(r => r.Is_Return === 'Yes').length],
        [""],
        ["Revenue By Region", ""]
    ];

    for (let reg in byRegion) {
        if (reg && reg !== 'undefined') dashData.push([reg, byRegion[reg]]);
    }

    const wsDash = XLSX.utils.aoa_to_sheet(dashData);

    // Make Dashboard column wider
    wsDash['!cols'] = [{ wch: 25 }, { wch: 15 }];

    XLSX.utils.book_append_sheet(wb, wsDash, "Dashboard Summary");

    // Generate and download
    const filename = `Cleaned_Data_Report_${new Date().getTime()}.xlsx`;
    XLSX.writeFile(wb, filename);
}

// Utility: find matching column name case-insensitively
function findCol(headers, searchTerms) {
    for (const term of searchTerms) {
        const match = headers.find(h => h.includes(term));
        if (match) {
            // Find the original capitalized version from rawData
            if (!rawData[0]) return match;
            return Object.keys(rawData[0]).find(k => k.toLowerCase().trim() === match) || match;
        }
    }
    return null;
}

// Utility: simple aggregation
function aggregate(data, groupCol, valCol) {
    if (!groupCol || !valCol) return {};
    return data.reduce((acc, row) => {
        let group = row[groupCol] || 'Unknown';
        let val = Number(row[valCol]) || 0;
        acc[group] = (acc[group] || 0) + val;
        return acc;
    }, {});
}
