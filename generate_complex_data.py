import csv
import random
import datetime

# Configuration
NUM_ROWS = 50000
FILENAME = "complex_enterprise_data.csv"

# Seed data
customers = ["Acme Corp", "Globex GmbH", "Soylent Corp", "Initech AG", "Umbrella Corp", "Stark Industries", "Wayne Enterprises", "Oscorp", "Massive Dynamic", "Cyberdyne Systems"]
categories = ["Hardware", "Software", "Consulting", "Licensing", "Support", "Training", "Cloud Infrastructure"]
regions = ["EMEA", "NA", "APAC", "LATAM", "Global"]

def generate_messy_date(base_date):
    """Generates dates in various formats to simulate messy enterprise data."""
    formats = [
        "%Y-%m-%d",      # Standard ISO
        "%d.%m.%Y",      # German 
        "%m/%d/%Y",      # US
        "%d/%m/%Y",      # UK
        "%Y.%m.%d"       # Weird format
    ]
    fmt = random.choice(formats)
    
    # Introduce occasional total garbage or empty dates
    rand_val = random.random()
    if rand_val < 0.05:
        return "" # 5% missing dates
    elif rand_val < 0.08:
        return "TBD" # 3% text instead of date
        
    return base_date.strftime(fmt)

def generate_messy_currency():
    """Generates currency values with varying formats (EU, US, messy)."""
    base_val = random.uniform(10.0, 50000.0)
    
    # Introduce negative values (returns)
    if random.random() < 0.05:
        base_val = -base_val

    rand_val = random.random()
    if rand_val < 0.1:
        return "" # 10% missing revenue
    elif rand_val < 0.4:
        # German format: 1.234,56 €
        return "{:,.2f}".format(base_val).replace(",", "X").replace(".", ",").replace("X", ".") + " €"
    elif rand_val < 0.7:
        # US format: $1,234.56
        return "${:,.2f}".format(base_val)
    else:
        # Just numbers
        return "{:.2f}".format(base_val)

def generate_messy_quantity():
    if random.random() < 0.05:
        return "" # 5% missing
    if random.random() < 0.02:
        return "N/A" # 2% bad text
    qty = random.randint(1, 1000)
    if random.random() < 0.05:
        qty = -qty # 5% negative
    return str(qty)

def generate_messy_customer():
    cust = random.choice(customers)
    # Simulate inconsistent casing and typos
    rand_val = random.random()
    if rand_val < 0.1:
        return cust.lower()
    elif rand_val < 0.2:
        return cust.upper()
    elif rand_val < 0.3:
        if "GmbH" in cust:
            return cust.replace("GmbH", "gmbh")
        if "AG" in cust:
            return cust.replace("AG", "ag")
    return cust


# Generate Data
start_date = datetime.date(2023, 1, 1)

print(f"Generating {NUM_ROWS} rows of highly complex enterprise data...")

with open(FILENAME, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    
    # Header with weird casing and spacing to test auto-detection
    writer.writerow([" Transaction ID ", "Date of Sale", " Client Name  ", "Product Group", " Item Desc ", " Qty Sold ", " Total Sales Value ", "Sales Territory "])
    
    for i in range(1, NUM_ROWS + 1):
        # Progress indicator
        if i % 10000 == 0:
            print(f"Generated {i} rows...")
            
        current_date = start_date + datetime.timedelta(days=random.randint(0, 700))
        
        row = [
            f"TXN-{random.randint(100000, 999999)}",
            generate_messy_date(current_date),
            generate_messy_customer(),
            random.choice(categories) if random.random() > 0.1 else "", # 10% missing categories
            f"Product-{random.randint(100, 999)}",
            generate_messy_quantity(),
            generate_messy_currency(),
            random.choice(regions) if random.random() > 0.05 else "" # 5% missing regions
        ]
        
        # Introduce occasional completely empty rows or duplicate rows (1%)
        if random.random() < 0.01:
            writer.writerow(["", "", "", "", "", "", "", ""])
        elif random.random() < 0.01 and i > 1:
            # write row twice (duplicate)
            writer.writerow(row)
            writer.writerow(row)
        else:
            writer.writerow(row)

print(f"Successfully generated {FILENAME}")
