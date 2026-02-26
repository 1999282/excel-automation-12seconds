# Excel Automation — 3 Hours → 12 Seconds

![Hero Image](web/og-preview.png)

> **"Most people apply with a CV. I show up with an AI arsenal."**
>
> A client-side, privacy-first web application that cleans messy CSVs and generates interactive reports in 12 seconds. Built to prove that automation beats manual labor in data analytics.

## 🚀 Live Demo (No Install Required)

### 👉 **[Try the Web App Here](https://1999282.github.io/excel-automation-12seconds/web/)**

---

## 🔬 "Expert Audit" Features

This tool was designed not just to "work," but to pass a Senior Data Analyst's code review. Recent upgrades include:

1. **100% Client-Side Processing:** Your data never leaves your browser (high privacy).
2. **Before vs. After Diff:** A side-by-side table showing exactly what characters and formatting were fixed in real-time.
3. **Data Quality Score:** Mathematically calculated percentage measuring the messiness improvement.
4. **Resilient Charting:** Safely destroys and recreates Chart.js instances across multiple files.
5. **Dynamic Column Detection:** Fuzzymatches column headers so it works with *your* data, not just the sample data.

---

## 💻 Tech Stack

- **Web App:** Vanilla HTML/CSS/JS (Zero build step, maximum speed)
- **Data Parsing:** PapaParse (CSV) & SheetJS (Excel)
- **Visuals:** Chart.js (interactive) + Custom CSS Confetti
- **Python Backend (Optional):** Pandas, Openpyxl, Matplotlib (for CLI automation fans)

---

## 🛠️ The 6-Step Pipeline

1. **Load:** Drag & drop CSV/Excel handling auto-encoding.
2. **Detect:** Intelligently maps column aliases (e.g. `qty`, `amount` -> Quantity).
3. **Clean:** Drops duplicates, standardizes dates (e.g. `15.01.2024` -> `2024-01-15`), fixes currency strings.
4. **Analyze:** Calculates margins, flags returns, aggregates totals.
5. **Visualize:** Renders 5 Chart.js insight panels (Trend, Region, Product, Margins, Top Customers).
6. **Export:** Generates a multi-sheet `.xlsx` file summarizing everything.

---

## ⚙️ How to Run Locally

If you don't want to use the live web app, you can run the original Python version:

```bash
git clone https://github.com/1999282/excel-automation-12seconds.git
cd excel-automation-12seconds
pip install pandas openpyxl matplotlib
python automate_report.py
```

## 📜 License

MIT License - Free to use, modify, and distribute.

*Built by Deepam Shah • Operations & Strategy • [LinkedIn](https://www.linkedin.com/in/deepammshah/)*
