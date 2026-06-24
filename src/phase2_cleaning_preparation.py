
# PHASE 2 — Data Cleaning & Preparation
# Retail Chain Sales Analytics


import pandas as pd
import sqlite3
import os

# ── STEP 1: Paths ────────────────────────────────────────────
BASE_DIR = r"C:\Users\KIIT\OneDrive\Desktop\data_analytic\Sales Data Analysis and Reporting for a Retail Chain"
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH  = os.path.join(DATA_DIR, "retail_analytics.db")

# ── STEP 2: Load from Database ───────────────────────────────
print("=" * 55)
print("  PHASE 2 — Data Cleaning & Preparation")
print("=" * 55)

conn = sqlite3.connect(DB_PATH)

df = pd.read_sql_query("SELECT * FROM fact_transactions", conn)
print(f"\nLoaded {len(df):,} rows from database")

# ── STEP 3: INITIAL DATA AUDIT ───────────────────────────────
print("\n--- STEP 3: Initial Data Audit ---")
print(f"  Shape             : {df.shape}")
print(f"  Columns           : {df.columns.tolist()}")
print(f"\nData Types:")
print(df.dtypes)
print(f"\nMissing Values:")
print(df.isnull().sum())
print(f"\nDuplicate Rows     : {df.duplicated().sum()}")
print(f"\nBasic Stats:")
print(df['tran_amount'].describe())

# ── STEP 4: SQL CLEANING CHECKS ──────────────────────────────
print("\n--- STEP 4: SQL Cleaning Checks ---")
cursor = conn.cursor()

# Check 1: Nulls in fact table
r = cursor.execute("""
    SELECT
        SUM(CASE WHEN customer_id  IS NULL THEN 1 ELSE 0 END) AS null_customer,
        SUM(CASE WHEN trans_date   IS NULL THEN 1 ELSE 0 END) AS null_date,
        SUM(CASE WHEN tran_amount  IS NULL THEN 1 ELSE 0 END) AS null_amount
    FROM fact_transactions
""").fetchone()
print(f"  NULL customer_id  : {r[0]}")
print(f"  NULL trans_date   : {r[1]}")
print(f"  NULL tran_amount  : {r[2]}")

# Check 2: Negative or zero amounts
r2 = cursor.execute("""
    SELECT COUNT(*) FROM fact_transactions
    WHERE tran_amount <= 0
""").fetchone()[0]
print(f"  Zero/Neg amounts  : {r2}")

# Check 3: Duplicates
r3 = cursor.execute("""
    SELECT COUNT(*) FROM (
        SELECT customer_id, trans_date, tran_amount,
               COUNT(*) as cnt
        FROM fact_transactions
        GROUP BY customer_id, trans_date, tran_amount
        HAVING cnt > 1
    )
""").fetchone()[0]
print(f"  Duplicate combos  : {r3}")

# Check 4: Date range
r4 = cursor.execute("""
    SELECT MIN(trans_date), MAX(trans_date)
    FROM fact_transactions
""").fetchone()
print(f"  Date range        : {r4[0]} to {r4[1]}")

# Check 5: Amount distribution
r5 = cursor.execute("""
    SELECT
        MIN(tran_amount)                    AS min_amount,
        MAX(tran_amount)                    AS max_amount,
        ROUND(AVG(tran_amount), 2)          AS avg_amount,
        COUNT(CASE WHEN tran_amount < 10
              THEN 1 END)                   AS below_10,
        COUNT(CASE WHEN tran_amount > 100
              THEN 1 END)                   AS above_100
    FROM fact_transactions
""").fetchone()
print(f"  Min amount        : ${r5[0]}")
print(f"  Max amount        : ${r5[1]}")
print(f"  Avg amount        : ${r5[2]}")
print(f"  Below $10         : {r5[3]}")
print(f"  Above $100        : {r5[4]}")

# ── STEP 5: PYTHON CLEANING ───────────────────────────────────
print("\n--- STEP 5: Python Cleaning (pandas) ---")

# Fix date type
df['trans_date'] = pd.to_datetime(df['trans_date'])

# 5a: Missing values
before = len(df)
missing = df.isnull().sum().sum()
print(f"  Total missing values    : {missing}")

# 5b: Remove duplicates
df.drop_duplicates(
    subset=['customer_id', 'trans_date', 'tran_amount'],
    inplace=True
)
removed_dups = before - len(df)
print(f"  Duplicates removed      : {removed_dups}")

# 5c: Remove invalid amounts
before2 = len(df)
df = df[df['tran_amount'] > 0]
removed_neg = before2 - len(df)
print(f"  Invalid amounts removed : {removed_neg}")

# 5d: Outlier detection using IQR
Q1  = df['tran_amount'].quantile(0.25)
Q3  = df['tran_amount'].quantile(0.75)
IQR = Q3 - Q1
lower = Q1 - 1.5 * IQR
upper = Q3 + 1.5 * IQR
outliers = df[(df['tran_amount'] < lower) | (df['tran_amount'] > upper)]
print(f"  IQR Lower Bound         : ${lower:.2f}")
print(f"  IQR Upper Bound         : ${upper:.2f}")
print(f"  Outliers detected       : {len(outliers)}")
# Flag outliers but keep them
df['is_outlier'] = (
    (df['tran_amount'] < lower) | (df['tran_amount'] > upper)
).astype(int)

# 5e: Validate customer_id format (must start with 'CS')
invalid_ids = df[~df['customer_id'].str.startswith('CS')]
print(f"  Invalid customer IDs    : {len(invalid_ids)}")
df = df[df['customer_id'].str.startswith('CS')]

print(f"\n  Final clean rows        : {len(df):,}")

# ── STEP 6: FEATURE ENGINEERING ──────────────────────────────
print("\n--- STEP 6: Feature Engineering ---")

# Date features (recalculate cleanly)
df['trans_date']   = pd.to_datetime(df['trans_date'])
df['year']         = df['trans_date'].dt.year
df['month_num']    = df['trans_date'].dt.month
df['month_name']   = df['trans_date'].dt.strftime('%B')
df['quarter']      = df['trans_date'].dt.quarter
df['day_name']     = df['trans_date'].dt.strftime('%A')
df['day_of_week']  = df['trans_date'].dt.dayofweek
df['week_of_year'] = df['trans_date'].dt.isocalendar().week.astype(int)
df['year_month']   = df['trans_date'].dt.strftime('%Y-%m')
df['fiscal_period']= 'FY' + df['year'].astype(str) + '-Q' + df['quarter'].astype(str)

# Transaction value buckets
df['amount_bucket'] = pd.cut(
    df['tran_amount'],
    bins=[0, 25, 50, 75, 100, 105],
    labels=['$0-25', '$26-50', '$51-75', '$76-100', '$100+']
)

# Customer lifetime value (total spent per customer)
clv = df.groupby('customer_id')['tran_amount'].sum().rename('customer_ltv')
df  = df.merge(clv, on='customer_id', how='left')

# Customer transaction count
txn_count = df.groupby('customer_id')['txn_id'].count().rename('customer_txn_count')
df = df.merge(txn_count, on='customer_id', how='left')

# Customer segment based on LTV
df['customer_segment'] = pd.qcut(
    df['customer_ltv'].rank(method='first'),
    q=4,
    labels=['Bronze', 'Silver', 'Gold', 'Platinum']
)

# Is weekend flag
df['is_weekend'] = (df['day_of_week'] >= 5).astype(int)

# Response label
df['responded'] = df['response'].map({1: 'Yes', 0: 'No'})

print("  New columns added:")
new_cols = ['year_month', 'fiscal_period', 'amount_bucket',
            'customer_ltv', 'customer_txn_count',
            'customer_segment', 'is_weekend', 'responded', 'is_outlier']
for col in new_cols:
    print(f"    ✔ {col}")

# ── STEP 7: SAVE CLEANED DATA BACK TO DB ─────────────────────
print("\n--- STEP 7: Saving Cleaned Data ---")

df['trans_date']     = df['trans_date'].astype(str)
df['amount_bucket']  = df['amount_bucket'].astype(str)
df['customer_segment'] = df['customer_segment'].astype(str)

df.to_sql('fact_transactions_clean', conn,
          if_exists='replace', index=False)

conn.commit()

# Verify
count = cursor.execute(
    "SELECT COUNT(*) FROM fact_transactions_clean"
).fetchone()[0]
print(f"  Saved to 'fact_transactions_clean': {count:,} rows")

# ── STEP 8: FINAL SUMMARY REPORT ─────────────────────────────
print("\n--- STEP 8: Cleaned Data Summary ---")

summary = cursor.execute("""
    SELECT
        COUNT(*)                        AS total_rows,
        COUNT(DISTINCT customer_id)     AS unique_customers,
        MIN(trans_date)                 AS earliest_date,
        MAX(trans_date)                 AS latest_date,
        SUM(tran_amount)                AS total_revenue,
        ROUND(AVG(tran_amount),2)       AS avg_transaction,
        MIN(tran_amount)                AS min_transaction,
        MAX(tran_amount)                AS max_transaction
    FROM fact_transactions_clean
""").fetchone()

print(f"  Total rows          : {summary[0]:,}")
print(f"  Unique customers    : {summary[1]:,}")
print(f"  Date range          : {summary[2]} to {summary[3]}")
print(f"  Total revenue       : ${summary[4]:,}")
print(f"  Avg transaction     : ${summary[5]}")
print(f"  Min transaction     : ${summary[6]}")
print(f"  Max transaction     : ${summary[7]}")

print("\nRevenue by Year (cleaned data):")
yearly = cursor.execute("""
    SELECT year,
           COUNT(*)          AS transactions,
           SUM(tran_amount)  AS revenue,
           ROUND(AVG(tran_amount),2) AS avg_amt
    FROM fact_transactions_clean
    GROUP BY year ORDER BY year
""").fetchall()
print(f"  {'Year':<6} {'Transactions':>14} {'Revenue':>12} {'Avg':>8}")
print("  " + "-"*44)
for row in yearly:
    print(f"  {row[0]:<6} {row[1]:>14,} ${row[2]:>11,} ${row[3]:>7}")

print("\nCustomer Segments:")
segs = cursor.execute("""
    SELECT customer_segment,
           COUNT(DISTINCT customer_id) AS customers,
           SUM(tran_amount)            AS revenue
    FROM fact_transactions_clean
    GROUP BY customer_segment
    ORDER BY revenue DESC
""").fetchall()
for row in segs:
    print(f"  {row[0]:<12} : {row[1]:>5,} customers   ${row[2]:>10,} revenue")

conn.close()

print("\nPhase 2 Complete!")
print(f"   Clean table saved : fact_transactions_clean")
print(f"   Database          : {DB_PATH}")