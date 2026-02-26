# ⚡ Excel Report Automation — 3 Hours → 12 Seconds

> **What took 3 hours manually now runs in 12 seconds with Python.**

A Python automation tool that takes raw, messy Excel/CSV data and produces a clean, formatted, analysis-ready report — automatically.

---

## 🎯 What It Does

| Task | Manual Time | Automated Time |
|------|-------------|---------------|
| Clean duplicates & missing values | 30 min | 0.8 sec |
| Standardize formats (dates, currencies, text) | 45 min | 1.2 sec |
| Generate summary statistics | 20 min | 0.5 sec |
| Create pivot tables & aggregations | 40 min | 1.5 sec |
| Build visual charts (bar, line, pie) | 45 min | 3.0 sec |
| Export formatted Excel report | 15 min | 2.0 sec |
| **TOTAL** | **~3 hours** | **~12 seconds** |

---

## 🚀 Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run with sample data (included)
python automate_report.py

# 3. Run with YOUR data
python automate_report.py --input your_data.csv
```

**Output**: A fully formatted Excel report (`output/report_YYYY-MM-DD.xlsx`) with:

- ✅ Cleaned data sheet
- ✅ Summary statistics sheet
- ✅ Pivot tables sheet
- ✅ Charts (saved as PNG)

---

## 📁 Project Structure

```
excel-automation-12seconds/
├── automate_report.py        # Main script — run this
├── requirements.txt          # Dependencies
├── sample_data/
│   └── messy_sales_data.csv  # Sample messy data to test with
├── output/                   # Generated reports go here
└── README.md                 # You're reading this
```

---

## 🛠️ Tech Stack

- **Python 3.8+**
- **pandas** — Data cleaning & transformation
- **openpyxl** — Excel file creation with formatting
- **matplotlib** — Chart generation

---

## 📊 Sample Output

The script generates a multi-sheet Excel report:

1. **Clean Data** — Duplicates removed, formats standardized, nulls handled
2. **Summary** — Key metrics, totals, averages, counts by category
3. **Pivot Analysis** — Revenue by region, product performance, time trends

---

## 💡 How to Customize

Edit the `CONFIG` section at the top of `automate_report.py`:

```python
CONFIG = {
    "date_columns": ["order_date", "ship_date"],    # Which columns are dates
    "currency_columns": ["revenue", "cost"],          # Which columns are currency
    "category_column": "region",                      # Group-by column for pivots
    "value_column": "revenue",                        # Main metric column
}
```

---

## 🤝 Contributing

Found a way to make it faster? Better? PRs welcome.

---

## 📝 License

MIT — Use it, modify it, build on it. Free forever.

---

**Built by [Deepam Shah](https://www.linkedin.com/in/deepammshah/) — Operations & Strategy**
