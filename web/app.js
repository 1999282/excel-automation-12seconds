/**
 * Excel Automation Web App
 * ========================
 * Client-side CSV processing with data cleaning, charts, and Excel export.
 * No backend — everything runs in the browser.
 */

// ============================================================
// Global State
// ============================================================
let rawData = [];
let cleanData = [];
let cleaningReport = {};
let chartInstances = [];
let processingStartTime = 0;

// Chart color palette
const COLORS = {
    primary: ['#3b82f6', '#8b5cf6', '#ec4899', '#f97316', '#22c55e', '#06b6d4', '#eab308'],
    gradient: [
        'rgba(59, 130, 246, 0.8)', 'rgba(139, 92, 246, 0.8)', 'rgba(236, 72, 153, 0.8)',
        'rgba(249, 115, 22, 0.8)', 'rgba(34, 197, 94, 0.8)', 'rgba(6, 182, 212, 0.8)'
    ],
    borders: [
        'rgba(59, 130, 246, 1)', 'rgba(139, 92, 246, 1)', 'rgba(236, 72, 153, 1)',
        'rgba(249, 115, 22, 1)', 'rgba(34, 197, 94, 1)', 'rgba(6, 182, 212, 1)'
    ]
};

// Chart.js global defaults for dark theme
Chart.defaults.color = '#8888a0';
Chart.defaults.borderColor = 'rgba(42, 42, 58, 0.5)';
Chart.defaults.font.family = "'Inter', sans-serif";

// ============================================================
// Upload Handlers
// ============================================================
const uploadZone = document.getElementById('upload-zone');
const fileInput = document.getElementById('file-input');

uploadZone.addEventListener('click', () => fileInput.click());
uploadZone.addEventListener('dragover', (e) => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
});

function handleFile(file) {
    if (!file.name.match(/\.(csv|xlsx|xls)$/i)) {
        alert('Please upload a CSV or Excel file.');
        return;
    }

    processingStartTime = performance.now();
    showProcessing();

    if (file.name.match(/\.csv$/i)) {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                rawData = results.data;
                processData();
            },
            error: (err) => {
                alert('Error reading file: ' + err.message);
                resetApp();
            }
        });
    } else {
        // Excel file
        const reader = new FileReader();
        reader.onload = (e) => {
            const wb = XLSX.read(e.target.result, { type: 'array' });
            const firstSheet = wb.Sheets[wb.SheetNames[0]];
            rawData = XLSX.utils.sheet_to_json(firstSheet);
            processData();
        };
        reader.readAsArrayBuffer(file);
    }
}

// ============================================================
// Sample Data
// ============================================================
function loadSampleData() {
    processingStartTime = performance.now();
    showProcessing();

    const sampleCSV = `order_id,order_date,customer_name,region,product,category,quantity,unit_price,revenue,cost,ship_date,status
1001,2025-01-15,Müller GmbH,NRW,Widget Pro,Electronics,25,49.99,1249.75,625.00,2025-01-18,Delivered
1002,2025-01-15,  SCHMIDT AG  ,nrw,Gadget X,electronics,10,89.99,899.90,400.00,2025-01-20,Shipped
1003,2025-01-16,Fischer Corp,Bavaria,Sensor Unit,Industrial,50,24.50,1225.00,612.50,,Delivered
1004,2025-01-17,Weber GmbH,Hamburg,cable kit,Accessories,30,12.99,389.70,156.00,2025-01-22,Delivered
1005,01/18/2025,Koch Solutions,Berlin,Widget Pro,Electronics,15,49.99,749.85,375.00,2025-01-23,Shipped
1006,2025-01-19,  bauer tech  ,NRW,Gadget X,ELECTRONICS,-5,89.99,-449.95,0,2025-01-24,Returned
1007,2025-01-20,Müller GmbH,NRW,Sensor Unit,Industrial,20,24.50,490.00,245.00,2025-01-25,Delivered
1008,2025-01-15,Müller GmbH,NRW,Widget Pro,Electronics,25,49.99,1249.75,625.00,2025-01-18,Delivered
1009,19 Jan 2025,Schneider Fabrik,Bavaria,Cable Kit,Accessories,100,12.99,1299.00,520.00,2025-01-26,Delivered
1010,2025-01-20,,NRW,Widget Pro,Electronics,40,49.99,1999.60,800.00,,Pending
1011,2025-01-21,Wagner Ltd,Hamburg,Power Module,Components,8,149.99,1199.92,480.00,2025-01-28,Delivered
1012,2025-01-22,Hoffmann KG,Berlin,Gadget X,Electronics,35,89.99,3149.65,1260.00,2025-01-29,Delivered
1013,2025-01-23,Koch Solutions,Berlin,Sensor Unit,Industrial,60,24.50,1470.00,588.00,2025-01-30,Shipped
1014,2025-01-24,Fischer Corp,Bavaria,Widget Pro,Electronics,18,49.99,899.82,360.00,,Delivered
1015,2025-01-25,Braun Logistik,NRW,Cable Kit,accessories,200,12.99,2598.00,1040.00,2025-02-01,Delivered
1016,2025-01-26,Keller Handels,Hamburg,Power Module,Components,12,149.99,1799.88,720.00,2025-02-02,Delivered
1017,2025-01-27,Müller GmbH,NRW,Gadget X,Electronics,30,89.99,2699.70,1080.00,2025-02-03,Delivered
1018,2025-01-28,Richter Systems,Bavaria,Sensor Unit,Industrial,45,24.50,1102.50,441.00,2025-02-04,Shipped
1019,2025-01-29,Wagner Ltd,Hamburg,Widget Pro,Electronics,22,49.99,1099.78,440.00,2025-02-05,Delivered
1020,2025-01-30,Schneider Fabrik,Bavaria,Cable Kit,Accessories,80,12.99,1039.20,416.00,2025-02-06,Delivered
1021,2025-01-15,Müller GmbH,NRW,Widget Pro,Electronics,25,49.99,1249.75,625.00,2025-01-18,Delivered
1022,2025-02-01,Meyer Transport,NRW,Power Module,Components,5,149.99,749.95,300.00,,Pending
1023,2025-02-02,Wolf Components,Berlin,Gadget X,Electronics,28,89.99,2519.72,1008.00,2025-02-09,Delivered
1024,2025-02-03,Schäfer Elektro,Bavaria,Sensor Unit,Industrial,55,24.50,1347.50,539.00,2025-02-10,Delivered`;

    Papa.parse(sampleCSV, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
            rawData = results.data;
            processData();
        }
    });
}

function scrollToUpload() {
    document.getElementById('upload-section').scrollIntoView({ behavior: 'smooth' });
}

// ============================================================
// Data Processing Pipeline
// ============================================================
function processData() {
    const report = {
        originalRows: rawData.length,
        duplicatesRemoved: 0,
        missingFilled: 0,
        formatsFixed: 0,
        negativesRemoved: 0,
        blankNamesFixed: 0
    };

    // Step 1: Loading
    updateStep(1, 'active');
    let data = JSON.parse(JSON.stringify(rawData));
    updateStep(1, 'done');

    // Step 2: Remove duplicates
    updateStep(2, 'active');
    const seen = new Set();
    const deduped = [];
    data.forEach(row => {
        const key = JSON.stringify(row);
        if (!seen.has(key)) {
            seen.add(key);
            deduped.push(row);
        }
    });
    report.duplicatesRemoved = data.length - deduped.length;
    data = deduped;
    updateStep(2, 'done');

    // Step 3: Clean data
    updateStep(3, 'active');

    // Detect column names (flexible)
    const cols = detectColumns(data[0]);

    data.forEach(row => {
        // Clean customer names
        if (cols.customer && row[cols.customer]) {
            const original = row[cols.customer];
            row[cols.customer] = row[cols.customer].toString().trim().replace(/\s+/g, ' ');
            // Title case
            row[cols.customer] = row[cols.customer].replace(/\w\S*/g, (txt) =>
                txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()
            );
            // Fix GmbH, AG, KG, Ltd casing
            row[cols.customer] = row[cols.customer]
                .replace(/\bGmbh\b/g, 'GmbH')
                .replace(/\bAg\b/g, 'AG')
                .replace(/\bKg\b/g, 'KG')
                .replace(/\bLtd\b/g, 'Ltd');

            if (original !== row[cols.customer]) report.formatsFixed++;
        }

        // Fill blank names
        if (cols.customer && (!row[cols.customer] || row[cols.customer].trim() === '')) {
            row[cols.customer] = 'Unknown Customer';
            report.blankNamesFixed++;
        }

        // Standardize region
        if (cols.region && row[cols.region]) {
            row[cols.region] = row[cols.region].toString().trim().toUpperCase();
            // Common corrections
            const regionMap = { 'NRW': 'NRW', 'BAVARIA': 'BAVARIA', 'HAMBURG': 'HAMBURG', 'BERLIN': 'BERLIN', 'BADEN-WÜRTTEMBERG': 'BADEN-WÜRTTEMBERG', 'SAXONY': 'SAXONY', 'HESSEN': 'HESSEN' };
            if (regionMap[row[cols.region]]) {
                row[cols.region] = regionMap[row[cols.region]];
            }
        }

        // Standardize category
        if (cols.category && row[cols.category]) {
            row[cols.category] = row[cols.category].toString().trim();
            row[cols.category] = row[cols.category].charAt(0).toUpperCase() + row[cols.category].slice(1).toLowerCase();
        }

        // Standardize product
        if (cols.product && row[cols.product]) {
            row[cols.product] = row[cols.product].toString().trim();
            row[cols.product] = row[cols.product].replace(/\b\w/g, l => l.toUpperCase());
        }

        // Parse numbers
        if (cols.quantity) row[cols.quantity] = parseFloat(row[cols.quantity]) || 0;
        if (cols.unit_price) row[cols.unit_price] = parseFloat(row[cols.unit_price]) || 0;
        if (cols.revenue) row[cols.revenue] = parseFloat(row[cols.revenue]) || 0;
        if (cols.cost) row[cols.cost] = parseFloat(row[cols.cost]) || 0;
    });

    // Remove negative quantities (returns)
    if (cols.quantity) {
        const before = data.length;
        data = data.filter(row => parseFloat(row[cols.quantity]) >= 0);
        report.negativesRemoved = before - data.length;
    }

    // Calculate profit & margin
    if (cols.revenue && cols.cost) {
        data.forEach(row => {
            row.profit = (parseFloat(row[cols.revenue]) || 0) - (parseFloat(row[cols.cost]) || 0);
            row.margin_pct = row[cols.revenue] > 0 ? ((row.profit / parseFloat(row[cols.revenue])) * 100).toFixed(1) : 0;
        });
    }

    updateStep(3, 'done');

    // Step 4: Stats
    updateStep(4, 'active');
    updateStep(4, 'done');

    // Step 5: Charts
    updateStep(5, 'active');
    updateStep(5, 'done');

    // Step 6: Report
    updateStep(6, 'active');
    report.cleanRows = data.length;
    cleanData = data;
    cleaningReport = report;

    updateStep(6, 'done');

    // Show results after animation
    setTimeout(() => {
        const elapsed = ((performance.now() - processingStartTime) / 1000).toFixed(1);
        showResults(elapsed, cols);
    }, 600);
}

// ============================================================
// Column Detection (flexible naming)
// ============================================================
function detectColumns(sampleRow) {
    if (!sampleRow) return {};
    const keys = Object.keys(sampleRow);
    const find = (patterns) => keys.find(k => patterns.some(p => k.toLowerCase().includes(p)));

    return {
        order_id: find(['order_id', 'id', 'order']),
        order_date: find(['date', 'order_date', 'created']),
        customer: find(['customer', 'client', 'company', 'name']),
        region: find(['region', 'area', 'state', 'location', 'city']),
        product: find(['product', 'item', 'sku']),
        category: find(['category', 'type', 'group']),
        quantity: find(['quantity', 'qty', 'units', 'count']),
        unit_price: find(['unit_price', 'price', 'unit']),
        revenue: find(['revenue', 'total', 'amount', 'sales']),
        cost: find(['cost', 'expense', 'cogs']),
        status: find(['status', 'state'])
    };
}

// ============================================================
// UI Helpers
// ============================================================
function showProcessing() {
    document.getElementById('upload-section').style.display = 'none';
    document.getElementById('hero').style.display = 'none';
    document.getElementById('processing-section').style.display = 'block';
}

function updateStep(n, state) {
    const el = document.getElementById(`step-${n}`);
    if (!el) return;
    el.className = `step ${state}`;
    if (state === 'done') {
        el.querySelector('.step-icon').textContent = '✅';
    } else if (state === 'active') {
        el.querySelector('.step-icon').textContent = '⚡';
    }
}

function showResults(elapsed, cols) {
    document.getElementById('processing-section').style.display = 'none';
    document.getElementById('results-section').style.display = 'block';

    // Timer
    document.getElementById('timer-value').textContent = elapsed + 's';

    // Summary cards
    const totalRevenue = cleanData.reduce((s, r) => s + (parseFloat(r[cols.revenue]) || 0), 0);
    const totalProfit = cleanData.reduce((s, r) => s + (parseFloat(r.profit) || 0), 0);
    const avgMargin = cleanData.length > 0 ? (cleanData.reduce((s, r) => s + (parseFloat(r.margin_pct) || 0), 0) / cleanData.length).toFixed(1) : 0;
    const uniqueCustomers = new Set(cleanData.map(r => r[cols.customer])).size;

    const summaryHTML = `
        <div class="summary-card"><div class="card-label">Total Orders</div><div class="card-value">${cleanData.length}</div></div>
        <div class="summary-card"><div class="card-label">Total Revenue</div><div class="card-value" style="color:#22c55e">€${formatNum(totalRevenue)}</div></div>
        <div class="summary-card"><div class="card-label">Total Profit</div><div class="card-value" style="color:#3b82f6">€${formatNum(totalProfit)}</div></div>
        <div class="summary-card"><div class="card-label">Avg Margin</div><div class="card-value" style="color:#8b5cf6">${avgMargin}%</div></div>
        <div class="summary-card"><div class="card-label">Customers</div><div class="card-value" style="color:#f97316">${uniqueCustomers}</div></div>
    `;
    document.getElementById('summary-grid').innerHTML = summaryHTML;

    // Cleaning report
    const reportHTML = `
        <div class="report-item"><span class="label">Original Rows</span><span class="value">${cleaningReport.originalRows}</span></div>
        <div class="report-item"><span class="label">Clean Rows</span><span class="value" style="color:#22c55e">${cleaningReport.cleanRows}</span></div>
        <div class="report-item"><span class="label">Duplicates Removed</span><span class="value" style="color:#ec4899">${cleaningReport.duplicatesRemoved}</span></div>
        <div class="report-item"><span class="label">Returns Removed</span><span class="value" style="color:#f97316">${cleaningReport.negativesRemoved}</span></div>
        <div class="report-item"><span class="label">Formats Fixed</span><span class="value" style="color:#8b5cf6">${cleaningReport.formatsFixed}</span></div>
        <div class="report-item"><span class="label">Blank Names Filled</span><span class="value" style="color:#06b6d4">${cleaningReport.blankNamesFixed}</span></div>
    `;
    document.getElementById('report-grid').innerHTML = reportHTML;

    // Build charts
    destroyCharts();
    buildCharts(cols);

    // Data table
    buildTable(cols);

    // Scroll to results
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
}

function formatNum(n) {
    return n.toLocaleString('de-DE', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

// ============================================================
// Charts
// ============================================================
function destroyCharts() {
    chartInstances.forEach(c => c.destroy());
    chartInstances = [];
}

function buildCharts(cols) {
    // 1. Revenue by Region
    if (cols.region && cols.revenue) {
        const grouped = groupBy(cleanData, cols.region, cols.revenue);
        const sorted = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
        chartInstances.push(new Chart(document.getElementById('chart-region'), {
            type: 'bar',
            data: {
                labels: sorted.map(s => s[0]),
                datasets: [{
                    label: 'Revenue (€)',
                    data: sorted.map(s => s[1]),
                    backgroundColor: COLORS.gradient,
                    borderColor: COLORS.borders,
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { color: 'rgba(42,42,58,0.3)' } },
                    y: { grid: { display: false } }
                }
            }
        }));
    }

    // 2. Revenue by Product (pie)
    if (cols.product && cols.revenue) {
        const grouped = groupBy(cleanData, cols.product, cols.revenue);
        const sorted = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
        chartInstances.push(new Chart(document.getElementById('chart-product'), {
            type: 'doughnut',
            data: {
                labels: sorted.map(s => s[0]),
                datasets: [{
                    data: sorted.map(s => s[1]),
                    backgroundColor: COLORS.gradient,
                    borderColor: '#16161f',
                    borderWidth: 3
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom', labels: { padding: 16, usePointStyle: true } }
                }
            }
        }));
    }

    // 3. Daily Trend (line)
    if (cols.order_date && cols.revenue) {
        const dailyData = {};
        cleanData.forEach(row => {
            let dateStr = row[cols.order_date];
            if (!dateStr) return;
            // Try to normalize the date
            let d = new Date(dateStr);
            if (isNaN(d.getTime())) {
                // Try DD/MM/YYYY
                const parts = dateStr.toString().split(/[\/\-\.]/);
                if (parts.length === 3) {
                    d = new Date(parts[2], parts[1] - 1, parts[0]);
                }
            }
            if (isNaN(d.getTime())) return;
            const key = d.toISOString().split('T')[0];
            dailyData[key] = (dailyData[key] || 0) + (parseFloat(row[cols.revenue]) || 0);
        });
        const sortedDays = Object.entries(dailyData).sort((a, b) => a[0].localeCompare(b[0]));

        chartInstances.push(new Chart(document.getElementById('chart-trend'), {
            type: 'line',
            data: {
                labels: sortedDays.map(d => d[0]),
                datasets: [{
                    label: 'Revenue (€)',
                    data: sortedDays.map(d => d[1]),
                    borderColor: '#3b82f6',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    fill: true,
                    tension: 0.3,
                    pointBackgroundColor: '#3b82f6',
                    pointBorderColor: '#fff',
                    pointRadius: 4,
                    pointHoverRadius: 6
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { color: 'rgba(42,42,58,0.3)' } },
                    y: { grid: { color: 'rgba(42,42,58,0.3)' } }
                }
            }
        }));
    }

    // 4. Top 5 Customers
    if (cols.customer && cols.revenue) {
        const grouped = groupBy(cleanData, cols.customer, cols.revenue);
        const sorted = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 5);
        chartInstances.push(new Chart(document.getElementById('chart-customers'), {
            type: 'bar',
            data: {
                labels: sorted.map(s => s[0]),
                datasets: [{
                    label: 'Revenue (€)',
                    data: sorted.map(s => s[1]),
                    backgroundColor: 'rgba(139, 92, 246, 0.7)',
                    borderColor: 'rgba(139, 92, 246, 1)',
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { color: 'rgba(42,42,58,0.3)' } },
                    y: { grid: { display: false } }
                }
            }
        }));
    }

    // 5. Profit Margins by Product
    if (cols.product && cleanData[0] && cleanData[0].margin_pct !== undefined) {
        const productMargins = {};
        const productCounts = {};
        cleanData.forEach(row => {
            const prod = row[cols.product];
            if (!prod) return;
            productMargins[prod] = (productMargins[prod] || 0) + parseFloat(row.margin_pct || 0);
            productCounts[prod] = (productCounts[prod] || 0) + 1;
        });
        const avgMargins = Object.entries(productMargins).map(([k, v]) => ([k, (v / productCounts[k]).toFixed(1)]));
        avgMargins.sort((a, b) => b[1] - a[1]);

        chartInstances.push(new Chart(document.getElementById('chart-margins'), {
            type: 'bar',
            data: {
                labels: avgMargins.map(m => m[0]),
                datasets: [{
                    label: 'Avg Margin %',
                    data: avgMargins.map(m => m[1]),
                    backgroundColor: avgMargins.map((_, i) => COLORS.gradient[i % COLORS.gradient.length]),
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { display: false } },
                    y: { grid: { color: 'rgba(42,42,58,0.3)' }, ticks: { callback: v => v + '%' } }
                }
            }
        }));
    }
}

function groupBy(data, groupCol, sumCol) {
    const result = {};
    data.forEach(row => {
        const key = row[groupCol] || 'Unknown';
        result[key] = (result[key] || 0) + (parseFloat(row[sumCol]) || 0);
    });
    return result;
}

// ============================================================
// Data Table
// ============================================================
function buildTable(cols) {
    const columns = Object.keys(cleanData[0] || {});
    const maxRows = Math.min(cleanData.length, 50);

    document.getElementById('table-count').textContent = `Showing ${maxRows} of ${cleanData.length} rows`;

    const thead = document.getElementById('table-head');
    thead.innerHTML = '<tr>' + columns.map(c => `<th>${c}</th>`).join('') + '</tr>';

    const tbody = document.getElementById('table-body');
    tbody.innerHTML = cleanData.slice(0, maxRows).map(row =>
        '<tr>' + columns.map(c => {
            let val = row[c];
            if (typeof val === 'number') val = val.toLocaleString('de-DE', { maximumFractionDigits: 2 });
            return `<td>${val ?? ''}</td>`;
        }).join('') + '</tr>'
    ).join('');
}

// ============================================================
// Excel Export
// ============================================================
function exportToExcel() {
    const wb = XLSX.utils.book_new();

    // Dashboard sheet
    const dashData = [
        ['EXCEL AUTOMATION REPORT'],
        ['Generated', new Date().toLocaleString()],
        [''],
        ['SUMMARY'],
        ['Total Orders', cleanData.length],
        ['Total Revenue', cleanData.reduce((s, r) => s + (parseFloat(r.revenue || r.Revenue || r.total || 0)), 0)],
        ['Total Profit', cleanData.reduce((s, r) => s + (parseFloat(r.profit) || 0), 0)],
        [''],
        ['CLEANING REPORT'],
        ['Original Rows', cleaningReport.originalRows],
        ['Clean Rows', cleaningReport.cleanRows],
        ['Duplicates Removed', cleaningReport.duplicatesRemoved],
        ['Returns Removed', cleaningReport.negativesRemoved],
    ];
    const wsDash = XLSX.utils.aoa_to_sheet(dashData);
    XLSX.utils.book_append_sheet(wb, wsDash, 'Dashboard');

    // Clean Data sheet
    const wsClean = XLSX.utils.json_to_sheet(cleanData);
    XLSX.utils.book_append_sheet(wb, wsClean, 'Clean Data');

    // Download
    XLSX.writeFile(wb, `report_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ============================================================
// Reset
// ============================================================
function resetApp() {
    rawData = [];
    cleanData = [];
    cleaningReport = {};
    destroyCharts();

    document.getElementById('results-section').style.display = 'none';
    document.getElementById('processing-section').style.display = 'none';
    document.getElementById('hero').style.display = '';
    document.getElementById('upload-section').style.display = '';

    // Reset steps
    for (let i = 1; i <= 6; i++) {
        const el = document.getElementById(`step-${i}`);
        if (el) {
            el.className = 'step';
            el.querySelector('.step-icon').textContent = '⏳';
        }
    }

    window.scrollTo({ top: 0, behavior: 'smooth' });
}
