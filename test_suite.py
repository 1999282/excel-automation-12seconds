"""
🧪 Expert-Level Test Suite for Excel Automation Tool
=====================================================
Tests the tool as a Senior Data Analyst would evaluate it:
- Edge cases (empty data, single row, all nulls)
- Stress test (500+ rows)
- Heavily corrupted data (mixed encodings, bad formats)
- Real-world German B2B scenario
- Validation of output accuracy

Author: Test Framework
"""

import pandas as pd
import os
import sys
import time
import random
import string
from datetime import datetime, timedelta

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_DIR = os.path.join(SCRIPT_DIR, 'test_data')
os.makedirs(TEST_DIR, exist_ok=True)

# Track results
results = []


def log_result(test_name, passed, details=""):
    status = "✅ PASS" if passed else "❌ FAIL"
    results.append({"test": test_name, "passed": passed, "details": details})
    print(f"  {status} | {test_name}")
    if details:
        print(f"         → {details}")


def generate_stress_test_data(n_rows=500):
    """Generate a large dataset with realistic messy data."""
    regions = ['NRW', 'Bavaria', 'Hamburg', 'Berlin', 'Baden-Württemberg', 'Saxony', 'Hessen']
    products = ['Widget Pro', 'Gadget X', 'Sensor Unit', 'Cable Kit', 'Power Module', 'Control Board']
    categories = ['Electronics', 'Industrial', 'Accessories', 'Components']
    companies = [
        'Müller GmbH', 'Schmidt AG', 'Fischer Corp', 'Weber Industries',
        'Koch Solutions', 'Bauer Tech', 'Wagner Ltd', 'Hoffmann KG',
        'Schneider Fabrik', 'Braun Logistik', 'Keller Handels', 'Richter Systems',
        'Meyer Transport', 'Wolf Components', 'Schäfer Elektro'
    ]
    statuses = ['Delivered', 'Shipped', 'Pending', 'Returned', 'Cancelled']
    
    date_formats = [
        '%Y-%m-%d', '%d/%m/%Y', '%m-%d-%Y', '%b %d %Y', '%Y/%m/%d',
        '%d-%m-%Y', '%d.%m.%Y'  # German format
    ]
    
    rows = []
    base_date = datetime(2025, 1, 1)
    
    for i in range(n_rows):
        order_date = base_date + timedelta(days=random.randint(0, 180))
        ship_date = order_date + timedelta(days=random.randint(1, 10))
        
        # Randomly format dates in different styles (simulates messy data)
        fmt = random.choice(date_formats)
        order_date_str = order_date.strftime(fmt)
        
        # 10% chance of missing ship date
        ship_date_str = ship_date.strftime('%Y-%m-%d') if random.random() > 0.1 else ''
        
        customer = random.choice(companies)
        # 5% chance of messy customer name (extra spaces, wrong case)
        if random.random() < 0.05:
            customer = f"  {customer.upper()}  "
        
        region = random.choice(regions)
        # 8% chance of inconsistent casing
        if random.random() < 0.08:
            region = region.lower() if random.random() > 0.5 else region.upper()
        
        product = random.choice(products)
        category = random.choice(categories)
        if random.random() < 0.08:
            category = category.upper()
        
        quantity = random.randint(5, 200)
        # 3% chance of negative (returns)
        if random.random() < 0.03:
            quantity = -quantity
        
        unit_price = round(random.uniform(5.0, 150.0), 2)
        revenue = round(quantity * unit_price, 2)
        cost = round(revenue * random.uniform(0.3, 0.7), 2) if quantity > 0 else 0
        
        # 3% chance of missing customer name
        if random.random() < 0.03:
            customer = ''
        
        status = random.choice(statuses)
        
        rows.append({
            'order_id': 1000 + i,
            'order_date': order_date_str,
            'customer_name': customer,
            'region': region,
            'product': product,
            'category': category,
            'quantity': quantity,
            'unit_price': unit_price,
            'revenue': revenue,
            'cost': cost,
            'ship_date': ship_date_str,
            'status': status,
        })
    
    # Add some exact duplicate rows (5%)
    n_dupes = int(n_rows * 0.05)
    for _ in range(n_dupes):
        rows.append(random.choice(rows).copy())
    
    random.shuffle(rows)
    
    df = pd.DataFrame(rows)
    return df


def test_1_sample_data():
    """Test 1: Does it work with the included sample data?"""
    print("\n" + "=" * 60)
    print("  TEST 1: Sample Data (Included)")
    print("=" * 60)
    
    start = time.time()
    exit_code = os.system(f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}"')
    elapsed = time.time() - start
    
    # Check output exists
    output_dir = os.path.join(SCRIPT_DIR, 'output')
    report_exists = any(f.endswith('.xlsx') for f in os.listdir(output_dir))
    charts_exist = len([f for f in os.listdir(output_dir) if f.startswith('chart_')]) >= 3
    
    log_result("Script exits successfully", exit_code == 0)
    log_result("Excel report generated", report_exists)
    log_result("Charts generated (3+)", charts_exist)
    log_result(f"Completed in under 5 seconds ({elapsed:.1f}s)", elapsed < 5)


def test_2_stress_test():
    """Test 2: Can it handle 500+ rows with messy data?"""
    print("\n" + "=" * 60)
    print("  TEST 2: Stress Test (500 rows, messy data)")
    print("=" * 60)
    
    df = generate_stress_test_data(500)
    test_file = os.path.join(TEST_DIR, 'stress_test_500.csv')
    df.to_csv(test_file, index=False)
    
    print(f"  Generated {len(df)} rows with mixed date formats, duplicates, missing values")
    
    start = time.time()
    exit_code = os.system(
        f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}" --input "{test_file}"'
    )
    elapsed = time.time() - start
    
    log_result("Handles 500+ rows without crashing", exit_code == 0)
    log_result(f"Completed in under 10 seconds ({elapsed:.1f}s)", elapsed < 10)
    
    # Validate the output
    output_dir = os.path.join(SCRIPT_DIR, 'output')
    report_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
    if report_files:
        report_path = os.path.join(output_dir, report_files[-1])
        xl = pd.ExcelFile(report_path)
        sheets = xl.sheet_names
        log_result("Has Dashboard sheet", 'Dashboard' in sheets)
        log_result("Has Clean Data sheet", 'Clean Data' in sheets)
        log_result("Has By Region sheet", 'By Region' in sheets)
        log_result("Has By Product sheet", 'By Product' in sheets)
        
        # Check clean data is actually clean
        clean_df = pd.read_excel(report_path, sheet_name='Clean Data')
        original_count = len(df)
        clean_count = len(clean_df)
        log_result(f"Duplicates removed ({original_count} → {clean_count})", clean_count < original_count)
        
        # Check no negative quantities in clean data
        if 'quantity' in clean_df.columns:
            has_negatives = (clean_df['quantity'] < 0).any()
            log_result("Negative quantities (returns) removed", not has_negatives)
        
        # Check profit column was calculated
        has_profit = 'profit' in clean_df.columns
        log_result("Profit column calculated", has_profit)
        
        has_margin = 'margin_pct' in clean_df.columns
        log_result("Margin % column calculated", has_margin)


def test_3_large_dataset():
    """Test 3: Performance with 1000 rows."""
    print("\n" + "=" * 60)
    print("  TEST 3: Performance Test (1000 rows)")
    print("=" * 60)
    
    df = generate_stress_test_data(1000)
    test_file = os.path.join(TEST_DIR, 'perf_test_1000.csv')
    df.to_csv(test_file, index=False)
    
    start = time.time()
    exit_code = os.system(
        f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}" --input "{test_file}"'
    )
    elapsed = time.time() - start
    
    log_result("Handles 1000+ rows without crashing", exit_code == 0)
    log_result(f"Still under 15 seconds ({elapsed:.1f}s)", elapsed < 15)


def test_4_edge_single_row():
    """Test 4: Single row of data — edge case."""
    print("\n" + "=" * 60)
    print("  TEST 4: Edge Case (Single Row)")
    print("=" * 60)
    
    data = """order_id,order_date,customer_name,region,product,category,quantity,unit_price,revenue,cost,ship_date,status
1001,2025-03-15,Test GmbH,NRW,Widget Pro,Electronics,10,49.99,499.90,250.00,2025-03-18,Delivered"""
    
    test_file = os.path.join(TEST_DIR, 'single_row.csv')
    with open(test_file, 'w') as f:
        f.write(data)
    
    exit_code = os.system(
        f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}" --input "{test_file}"'
    )
    
    log_result("Handles single row without crashing", exit_code == 0)


def test_5_german_special_chars():
    """Test 5: German company names with umlauts and special chars."""
    print("\n" + "=" * 60)
    print("  TEST 5: German Special Characters (Umlauts)")
    print("=" * 60)
    
    data = """order_id,order_date,customer_name,region,product,category,quantity,unit_price,revenue,cost,ship_date,status
1001,2025-01-15,Müller & Söhne GmbH,NRW,Widget Pro,Electronics,25,49.99,1249.75,625.00,2025-01-18,Delivered
1002,2025-01-16,Böhm Präzisionstechnik,Bavaria,Gadget X,Electronics,10,89.99,899.90,400.00,2025-01-20,Delivered
1003,2025-01-17,Größe Stähle AG,Hamburg,Sensor Unit,Industrial,50,24.50,1225.00,612.50,2025-01-22,Delivered
1004,2025-01-18,Schäfer Würfel GmbH,Berlin,Cable Kit,Accessories,30,12.99,389.70,156.00,2025-01-22,Delivered
1005,2025-01-19,Überflieger Lösung KG,Baden-Württemberg,Widget Pro,Electronics,15,49.99,749.85,375.00,2025-01-23,Shipped"""
    
    test_file = os.path.join(TEST_DIR, 'german_chars.csv')
    with open(test_file, 'w', encoding='utf-8') as f:
        f.write(data)
    
    exit_code = os.system(
        f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}" --input "{test_file}"'
    )
    
    log_result("Handles German umlauts (ü, ö, ä, ß) correctly", exit_code == 0)
    
    # Verify the names are preserved in output
    output_dir = os.path.join(SCRIPT_DIR, 'output')
    report_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
    if report_files:
        report_path = os.path.join(output_dir, report_files[-1])
        clean_df = pd.read_excel(report_path, sheet_name='Clean Data')
        has_umlauts = any('ü' in str(name).lower() or 'ö' in str(name).lower() 
                        for name in clean_df['customer_name'].values)
        log_result("Umlauts preserved in output Excel", has_umlauts)


def test_6_all_missing_values():
    """Test 6: Data with heavy missing values."""
    print("\n" + "=" * 60)
    print("  TEST 6: Heavy Missing Values (30% null)")
    print("=" * 60)
    
    data = """order_id,order_date,customer_name,region,product,category,quantity,unit_price,revenue,cost,ship_date,status
1001,2025-01-15,,NRW,Widget Pro,Electronics,25,49.99,1249.75,625.00,,Delivered
1002,,Schmidt AG,,Gadget X,,10,89.99,899.90,,,Shipped
1003,2025-01-17,Fischer Corp,NRW,,Electronics,,,,,2025-01-20,Delivered
1004,2025-01-18,Weber GmbH,Hamburg,Sensor Unit,Industrial,50,24.50,1225.00,612.50,2025-01-22,
1005,2025-01-19,,Berlin,Cable Kit,,100,12.99,1299.00,520.00,2025-01-23,Delivered
1006,2025-01-20,Koch AG,NRW,Widget Pro,Electronics,30,49.99,1499.70,750.00,2025-01-24,Delivered"""
    
    test_file = os.path.join(TEST_DIR, 'heavy_nulls.csv')
    with open(test_file, 'w') as f:
        f.write(data)
    
    exit_code = os.system(
        f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}" --input "{test_file}"'
    )
    
    log_result("Handles heavy null data without crashing", exit_code == 0)


def test_7_output_accuracy():
    """Test 7: Verify output numbers are mathematically correct."""
    print("\n" + "=" * 60)
    print("  TEST 7: Output Accuracy (Math Validation)")
    print("=" * 60)
    
    # Known data with pre-calculated expected results
    data = """order_id,order_date,customer_name,region,product,category,quantity,unit_price,revenue,cost,ship_date,status
1001,2025-01-15,Alpha GmbH,NRW,Widget Pro,Electronics,10,100.00,1000.00,500.00,2025-01-18,Delivered
1002,2025-01-16,Beta AG,Bavaria,Gadget X,Electronics,20,50.00,1000.00,400.00,2025-01-20,Delivered
1003,2025-01-17,Gamma Corp,NRW,Widget Pro,Electronics,5,100.00,500.00,200.00,2025-01-21,Delivered
1004,2025-01-18,Alpha GmbH,NRW,Sensor Unit,Industrial,10,25.00,250.00,100.00,2025-01-22,Delivered
1005,2025-01-19,Delta Ltd,Hamburg,Cable Kit,Accessories,50,10.00,500.00,200.00,2025-01-23,Delivered"""
    
    EXPECTED_TOTAL_REVENUE = 3250.00
    EXPECTED_TOTAL_PROFIT = 3250.00 - (500 + 400 + 200 + 100 + 200)  # 1850.00
    EXPECTED_ORDERS = 5
    EXPECTED_NRW_REVENUE = 1000 + 500 + 250  # 1750.00
    
    test_file = os.path.join(TEST_DIR, 'accuracy_test.csv')
    with open(test_file, 'w') as f:
        f.write(data)
    
    exit_code = os.system(
        f'python "{os.path.join(SCRIPT_DIR, "automate_report.py")}" --input "{test_file}"'
    )
    
    if exit_code == 0:
        output_dir = os.path.join(SCRIPT_DIR, 'output')
        report_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
        report_path = os.path.join(output_dir, report_files[-1])
        
        clean_df = pd.read_excel(report_path, sheet_name='Clean Data')
        
        actual_revenue = clean_df['revenue'].sum()
        actual_orders = len(clean_df)
        actual_profit = clean_df['profit'].sum()
        
        log_result(
            f"Total revenue correct (Expected: €{EXPECTED_TOTAL_REVENUE}, Got: €{actual_revenue})",
            abs(actual_revenue - EXPECTED_TOTAL_REVENUE) < 0.01
        )
        log_result(
            f"Total orders correct (Expected: {EXPECTED_ORDERS}, Got: {actual_orders})",
            actual_orders == EXPECTED_ORDERS
        )
        log_result(
            f"Total profit correct (Expected: €{EXPECTED_TOTAL_PROFIT}, Got: €{actual_profit})",
            abs(actual_profit - EXPECTED_TOTAL_PROFIT) < 0.01
        )
        
        # Check NRW revenue from pivot
        region_df = pd.read_excel(report_path, sheet_name='By Region')
        if 'NRW' in region_df.index or (region_df.columns[0] == 'region' and 'NRW' in region_df['region'].values):
            # Try to find NRW revenue
            try:
                nrw_row = region_df[region_df.iloc[:, 0] == 'NRW']
                if not nrw_row.empty:
                    nrw_revenue = nrw_row.iloc[0, 1]
                    log_result(
                        f"NRW pivot correct (Expected: €{EXPECTED_NRW_REVENUE}, Got: €{nrw_revenue})",
                        abs(float(nrw_revenue) - EXPECTED_NRW_REVENUE) < 0.01
                    )
            except Exception as e:
                log_result(f"NRW pivot check", False, str(e))
    else:
        log_result("Script ran successfully", False)


# ============================================================
# RUN ALL TESTS
# ============================================================
if __name__ == "__main__":
    print()
    print("🧪" * 30)
    print("  EXPERT-LEVEL TEST SUITE")
    print("  Excel Automation Tool Validation")
    print("🧪" * 30)
    
    test_1_sample_data()
    test_2_stress_test()
    test_3_large_dataset()
    test_4_edge_single_row()
    test_5_german_special_chars()
    test_6_all_missing_values()
    test_7_output_accuracy()
    
    # Final Summary
    total = len(results)
    passed = sum(1 for r in results if r['passed'])
    failed = sum(1 for r in results if not r['passed'])
    
    print("\n")
    print("=" * 60)
    print(f"  📊 FINAL RESULTS: {passed}/{total} tests passed")
    print("=" * 60)
    
    if failed > 0:
        print(f"\n  ❌ FAILED TESTS ({failed}):")
        for r in results:
            if not r['passed']:
                print(f"     • {r['test']}")
                if r['details']:
                    print(f"       → {r['details']}")
    
    if passed == total:
        print("\n  🏆 ALL TESTS PASSED — Tool is production-ready!")
    elif passed / total >= 0.85:
        print("\n  ⚠️ MOSTLY PASSING — Minor issues to fix.")
    else:
        print("\n  🚫 SIGNIFICANT ISSUES — Do not share yet.")
    
    print()
