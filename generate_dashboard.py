#!/usr/bin/env python3
"""
FY26 Dashboard Generator
Extracts data from Excel files and generates standalone HTML dashboard
"""

import pandas as pd
import json
from datetime import datetime
from collections import defaultdict
import re

# ============ CONFIG ============
SALES_FILES = {
    'main': 'final sales register.xlsx',
    'guj': 'final sales register.xlsx'  # Update if separate file exists
}
BILL_FILE = 'Bill Wise Detail Final 25-26.xlsx'
PURCHASE_FILE = 'PURCHASE DATA - 2025-26.xls'

MONTHS_ORDER = ["2024-04","2024-05","2024-06","2024-07","2024-08","2024-09","2024-10","2024-11","2024-12","2025-01","2025-02","2025-03"]
MONTH_LABELS = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]

PARTY_GROUPS = {
    'Reliance': 'Reliance',
    'Lifestyle': 'Lifestyle',
    'Nexon': 'Nexon',
    'D-Mart': 'D-Mart',
    'Amazon': 'Amazon/SmartPaddle',
    'SmartPaddle': 'Amazon/SmartPaddle',
    'Shoppers Stop': 'Shoppers Stop',
    'Mogli': 'Mogli Labs',
    'Myntra': 'Myntra',
}

# ============ HELPERS ============
def parse_date(d):
    """Parse date from various formats"""
    if pd.isna(d):
        return None
    if isinstance(d, str):
        try:
            return pd.to_datetime(d)
        except:
            return None
    return d

def get_month_key(date_obj):
    """Convert date to YYYY-MM format"""
    if date_obj is None:
        return None
    return f"{date_obj.year:04d}-{date_obj.month:02d}"

def normalize_company(name):
    """Normalize party name to standard grouping"""
    if pd.isna(name):
        return 'Others'
    name = str(name).strip().upper()
    for key, group in PARTY_GROUPS.items():
        if key.upper() in name:
            return group
    return 'Others'

def safe_float(val):
    """Safely convert to float"""
    try:
        if pd.isna(val):
            return 0.0
        return float(val)
    except:
        return 0.0

# ============ DATA EXTRACTION ============

def load_sales_data():
    """Load and combine sales register data"""
    print("📊 Loading sales data...")
    sales_data = []
    
    try:
        xls = pd.ExcelFile('final sales register.xlsx')
        for sheet in xls.sheet_names:
            if sheet.lower() in ['summary', 'info']:
                continue
            df = pd.read_excel('final sales register.xlsx', sheet_name=sheet)
            # Standardize column names
            df.columns = [c.strip().lower() for c in df.columns]
            sales_data.append(df)
            print(f"  ✓ Loaded sheet: {sheet} ({len(df)} rows)")
    except Exception as e:
        print(f"  ✗ Error loading sales data: {e}")
    
    return pd.concat(sales_data, ignore_index=True) if sales_data else pd.DataFrame()

def load_bill_wise_data():
    """Load bill wise detail data (exclude Summary, STITCHER)"""
    print("📦 Loading bill wise data...")
    bill_data = []
    
    try:
        xls = pd.ExcelFile('Bill Wise Detail Final 25-26.xlsx')
        for sheet in xls.sheet_names:
            if sheet.lower() in ['summary', 'stitcher']:
                continue
            df = pd.read_excel('Bill Wise Detail Final 25-26.xlsx', sheet_name=sheet)
            df.columns = [c.strip().lower() for c in df.columns]
            bill_data.append(df)
            print(f"  ✓ Loaded sheet: {sheet} ({len(df)} rows)")
    except Exception as e:
        print(f"  ✗ Error loading bill wise data: {e}")
    
    return pd.concat(bill_data, ignore_index=True) if bill_data else pd.DataFrame()

def load_purchase_data():
    """Load purchase data"""
    print("💰 Loading purchase data...")
    purchase_data = []
    
    try:
        xls = pd.ExcelFile('PURCHASE DATA - 2025-26.xls')
        for sheet in xls.sheet_names:
            df = pd.read_excel('PURCHASE DATA - 2025-26.xls', sheet_name=sheet)
            df.columns = [c.strip().lower() for c in df.columns]
            purchase_data.append(df)
            print(f"  ✓ Loaded sheet: {sheet} ({len(df)} rows)")
    except Exception as e:
        print(f"  ✗ Error loading purchase data: {e}")
    
    return pd.concat(purchase_data, ignore_index=True) if purchase_data else pd.DataFrame()

# ============ AGGREGATION ============

def aggregate_sales(df):
    """Aggregate sales metrics by month and customer"""
    print("📈 Aggregating sales data...")
    
    monthly_sales = defaultdict(float)
    monthly_co_sales = defaultdict(lambda: defaultdict(float))
    monthly_co_returns = defaultdict(lambda: defaultdict(float))
    monthly_co_pcs = defaultdict(lambda: defaultdict(float))
    co_sales = defaultdict(float)
    co_returns = defaultdict(float)
    co_pcs = defaultdict(float)
    co_dn1 = defaultdict(float)
    co_cn1 = defaultdict(float)
    co_return_pcs = defaultdict(float)
    
    for _, row in df.iterrows():
        # Skip GSR/OTR rows
        prefix = str(row.get('prefix', '')).upper() if 'prefix' in row else ''
        if prefix in ['GSR', 'OTR']:
            continue
        
        # Extract key fields
        invoice_date = parse_date(row.get('date') or row.get('invoice date'))
        if invoice_date is None:
            continue
        
        month_key = get_month_key(invoice_date)
        if month_key is None or month_key not in MONTHS_ORDER:
            continue
        
        customer = normalize_company(row.get('party') or row.get('customer'))
        gross_amt = safe_float(row.get('grossamt') or row.get('gross amount'))
        net_amt = safe_float(row.get('netamt') or row.get('net amount'))
        pcs = safe_float(row.get('pcs') or row.get('qty'))
        
        # Classify as sales or returns
        if prefix == 'GRS/':
            monthly_co_returns[month_key][customer] += net_amt
            co_returns[customer] += net_amt
            # Extract return pcs if available
            if pcs > 0:
                co_return_pcs[customer] += pcs
        else:
            # Regular sales
            monthly_sales[month_key] += gross_amt
            monthly_co_sales[month_key][customer] += gross_amt
            co_sales[customer] += gross_amt
            if pcs > 0:
                monthly_co_pcs[month_key][customer] += pcs
                co_pcs[customer] += pcs
        
        # DN1/CN1 adjustments
        doc_type = str(row.get('document type', '')).upper()
        if 'DN1' in doc_type or 'DEBIT' in doc_type:
            co_dn1[customer] += net_amt
        elif 'CN1' in doc_type or 'CREDIT' in doc_type:
            co_cn1[customer] += net_amt
    
    return {
        'monthly_sales': dict(monthly_sales),
        'monthly_co_sales': {k: dict(v) for k, v in monthly_co_sales.items()},
        'monthly_co_returns': {k: dict(v) for k, v in monthly_co_returns.items()},
        'monthly_co_pcs': {k: dict(v) for k, v in monthly_co_pcs.items()},
        'co_sales': dict(co_sales),
        'co_returns': dict(co_returns),
        'co_pcs': dict(co_pcs),
        'co_dn1': dict(co_dn1),
        'co_cn1': dict(co_cn1),
        'co_return_pcs': dict(co_return_pcs),
    }

def aggregate_bill_wise(df):
    """Aggregate bill wise metrics"""
    print("📦 Aggregating bill wise data...")
    
    bill_monthly_short = defaultdict(lambda: {'pcs': 0, 'grn': 0, 'short': 0})
    stitcher_data = defaultdict(lambda: {'amt': 0, 'bills': 0, 'pcs': 0, 'grn': 0, 'short': 0})
    stitch_vendor = defaultdict(float)
    
    for _, row in df.iterrows():
        month_str = str(row.get('month') or row.get('date'))[:7]
        if month_str not in MONTHS_ORDER:
            continue
        
        pcs = safe_float(row.get('dispatched') or row.get('pcs'))
        grn = safe_float(row.get('grn') or row.get('received'))
        stitcher = str(row.get('stitcher') or 'Unknown').strip()
        amt = safe_float(row.get('amount') or row.get('billed'))
        
        short = pcs - grn
        bill_monthly_short[month_str]['pcs'] += pcs
        bill_monthly_short[month_str]['grn'] += grn
        bill_monthly_short[month_str]['short'] += short
        
        stitcher_data[stitcher]['amt'] += amt
        stitcher_data[stitcher]['bills'] += 1
        stitcher_data[stitcher]['pcs'] += pcs
        stitcher_data[stitcher]['grn'] += grn
        stitcher_data[stitcher]['short'] += short
        
        stitch_vendor[stitcher] += amt
    
    return {
        'bill_monthly_short': {k: dict(v) for k, v in bill_monthly_short.items()},
        'stitcher_data': {k: dict(v) for k, v in stitcher_data.items()},
        'stitch_vendor': dict(stitch_vendor),
    }

def aggregate_purchase(df):
    """Aggregate purchase data by category"""
    print("💰 Aggregating purchase data...")
    
    monthly_purchase_total = defaultdict(float)
    monthly_purch_by_cat = defaultdict(lambda: defaultdict(float))
    monthly_margin = defaultdict(float)
    cat_totals = defaultdict(float)
    cat_units = {}
    
    for _, row in df.iterrows():
        month_str = str(row.get('month') or row.get('date'))[:7]
        if month_str not in MONTHS_ORDER:
            continue
        
        netamt = safe_float(row.get('netamt') or row.get('net amount'))
        category = str(row.get('category') or row.get('accled_name', '')).strip()
        unit = str(row.get('unit') or row.get('unitname', '')).strip()
        
        if category:
            monthly_purchase_total[month_str] += netamt
            monthly_purch_by_cat[month_str][category] += netamt
            cat_totals[category] += netamt
            if category not in cat_units:
                cat_units[category] = unit
    
    # Calculate margins
    monthly_sales = defaultdict(float)
    monthly_net = defaultdict(float)
    # (Would need to pass sales data to calculate margins properly)
    
    return {
        'monthly_purchase_total': dict(monthly_purchase_total),
        'monthly_purch_by_cat': {k: dict(v) for k, v in monthly_purch_by_cat.items()},
        'monthly_margin': dict(monthly_margin),
        'monthly_net': dict(monthly_net),
        'cat_totals': dict(cat_totals),
        'cat_units': cat_units,
    }

# ============ MAIN ============

def main():
    print("🚀 FY26 Dashboard Generator\n")
    
    # Load data
    sales_df = load_sales_data()
    bill_df = load_bill_wise_data()
    purchase_df = load_purchase_data()
    
    # Aggregate
    sales_agg = aggregate_sales(sales_df)
    bill_agg = aggregate_bill_wise(bill_df)
    purchase_agg = aggregate_purchase(purchase_df)
    
    # Combine
    combined_data = {
        'months_order': MONTHS_ORDER,
        'month_labels': MONTH_LABELS,
        'companies': list(set(
            list(sales_agg['co_sales'].keys()) + 
            list(sales_agg['co_returns'].keys())
        )),
        **sales_agg,
        **bill_agg,
        **purchase_agg,
    }
    
    # Calculate margins
    for month in MONTHS_ORDER:
        s = combined_data['monthly_sales'].get(month, 0)
        p = combined_data['monthly_purchase_total'].get(month, 0)
        combined_data['monthly_margin'][month] = s - p
        combined_data['monthly_net'][month] = s
    
    print(f"\n✅ Data aggregation complete!")
    print(f"   Companies: {len(combined_data['companies'])}")
    print(f"   Total Gross Sales: ₹{sum(combined_data['monthly_sales'].values()):,.0f}")
    print(f"   Total Purchase: ₹{sum(combined_data['monthly_purchase_total'].values()):,.0f}")
    
    # Generate JSON
    data_json = json.dumps(combined_data, separators=(',', ':'))
    print(f"   Data JSON size: {len(data_json):,} bytes")
    
    print("\n📝 Data generation complete. Save this to HTML template.")
    return combined_data

if __name__ == '__main__':
    data = main()
