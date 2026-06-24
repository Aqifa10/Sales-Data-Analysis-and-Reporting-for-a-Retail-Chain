# PHASE 1 — Data Collection & Database Setup

import pandas as pd
import sqlite3
import os

# ── STEP 1: Set Paths ────────────────────────────────────────
BASE_DIR = r"C:\Users\KIIT\OneDrive\Desktop\data_analytic\Sales Data Analysis and Reporting for a Retail Chain"

DATA_DIR  = os.path.join(BASE_DIR, "data")
DB_PATH   = os.path.join(DATA_DIR, "retail_analytics.db")
TX_CSV    = os.path.join(DATA_DIR, "Retail_Data_Transactions.csv")
RESP_CSV  = os.path.join(DATA_DIR, "Retail_Data_Response.csv")

# ── STEP 2: Load CSVs ────────────────────────────────────────
print("Loading CSVs...")

transactions = pd.read_csv(TX_CSV)
response     = pd.read_csv(RESP_CSV)

print(f"  Transactions loaded : {len(transactions):,} rows")
print(f"  Response loaded     : {len(response):,} rows")

# ── STEP 3: Basic Cleaning Before Load ───────────────────────
print("\nCleaning data...")

transactions['trans_date'] = pd.to_datetime(
    transactions['trans_date'], format='%d-%b-%y'
)

transactions['year']        = transactions['trans_date'].dt.year
transactions['month_num']   = transactions['trans_date'].dt.month
transactions['month_name']  = transactions['trans_date'].dt.strftime('%b')
transactions['quarter']     = transactions['trans_date'].dt.quarter
transactions['day_name']    = transactions['trans_date'].dt.strftime('%A')
transactions['day_of_week'] = transactions['trans_date'].dt.dayofweek

df = transactions.merge(response, on='customer_id', how='left')
df['response'] = df['response'].fillna(0).astype(int)

print(f"  Merged shape        : {df.shape}")
print(f"  Date range          : {df['trans_date'].min().date()} to {df['trans_date'].max().date()}")
print(f"  Unique customers    : {df['customer_id'].nunique():,}")
print(f"  Amount range        : ${df['tran_amount'].min()} to ${df['tran_amount'].max()}")

# ── STEP 4: Connect to SQLite ─────────────────────────────────
print("\nSetting up SQLite database...")
conn   = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

# ── STEP 5: Create Tables 
cursor.executescript("""
    DROP TABLE IF EXISTS fact_transactions;
    DROP TABLE IF EXISTS dim_customer;
    DROP TABLE IF EXISTS dim_date;
    DROP TABLE IF EXISTS dim_response;
""")

cursor.execute("""
    CREATE TABLE dim_customer (
        customer_id     TEXT PRIMARY KEY,
        total_spent     REAL,
        total_txns      INTEGER,
        avg_txn_amount  REAL,
        first_purchase  TEXT,
        last_purchase   TEXT
    )
""")

cursor.execute("""
    CREATE TABLE dim_date (
        date_key    TEXT PRIMARY KEY,
        year        INTEGER,
        month_num   INTEGER,
        month_name  TEXT,
        quarter     INTEGER,
        day_name    TEXT,
        day_of_week INTEGER
    )
""")

cursor.execute("""
    CREATE TABLE dim_response (
        customer_id TEXT PRIMARY KEY,
        response    INTEGER
    )
""")

cursor.execute("""
    CREATE TABLE fact_transactions (
        txn_id          INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_id     TEXT    NOT NULL,
        trans_date      TEXT    NOT NULL,
        tran_amount     INTEGER NOT NULL,
        year            INTEGER,
        month_num       INTEGER,
        month_name      TEXT,
        quarter         INTEGER,
        day_name        TEXT,
        day_of_week     INTEGER,
        response        INTEGER DEFAULT 0,
        FOREIGN KEY (customer_id) REFERENCES dim_customer(customer_id)
    )
""")

cursor.executescript("""
    CREATE INDEX idx_txn_customer ON fact_transactions(customer_id);
    CREATE INDEX idx_txn_date     ON fact_transactions(trans_date);
    CREATE INDEX idx_txn_year     ON fact_transactions(year);
    CREATE INDEX idx_txn_amount   ON fact_transactions(tran_amount);
""")

print("  Tables created: fact_transactions, dim_customer, dim_date, dim_response")

# ── STEP 6: Populate Dimension Tables ─────────────────────────
cust_summary = df.groupby('customer_id').agg(
    total_spent    = ('tran_amount', 'sum'),
    total_txns     = ('tran_amount', 'count'),
    avg_txn_amount = ('tran_amount', 'mean'),
    first_purchase = ('trans_date',  'min'),
    last_purchase  = ('trans_date',  'max')
).reset_index()
cust_summary['first_purchase']  = cust_summary['first_purchase'].astype(str)
cust_summary['last_purchase']   = cust_summary['last_purchase'].astype(str)
cust_summary['avg_txn_amount']  = cust_summary['avg_txn_amount'].round(2)
cust_summary.to_sql('dim_customer', conn, if_exists='append', index=False)

dates = df[['trans_date','year','month_num','month_name',
            'quarter','day_name','day_of_week']].copy()
dates['date_key'] = dates['trans_date'].astype(str)
dates = dates.drop_duplicates('date_key')[
    ['date_key','year','month_num','month_name','quarter','day_name','day_of_week']
]
dates.to_sql('dim_date', conn, if_exists='append', index=False)

response.to_sql('dim_response', conn, if_exists='append', index=False)

# ── STEP 7: Populate Fact Table ───────────────────────────────
fact = df[['customer_id','trans_date','tran_amount','year',
           'month_num','month_name','quarter','day_name',
           'day_of_week','response']].copy()
fact['trans_date'] = fact['trans_date'].astype(str)
fact.to_sql('fact_transactions', conn, if_exists='append', index=False)

conn.commit()

# ── STEP 8: Verify ────────────────────────────────────────────
print("\nVerifying database...")

checks = {
    "fact_transactions" : "SELECT COUNT(*) FROM fact_transactions",
    "dim_customer"      : "SELECT COUNT(*) FROM dim_customer",
    "dim_date"          : "SELECT COUNT(*) FROM dim_date",
    "dim_response"      : "SELECT COUNT(*) FROM dim_response",
}
for table, query in checks.items():
    count = cursor.execute(query).fetchone()[0]
    print(f"  {table:<22}: {count:,} rows")

print("\nSample query — Revenue by Year:")
result = cursor.execute("""
    SELECT year,
           COUNT(*)         AS transactions,
           SUM(tran_amount) AS total_revenue,
           ROUND(AVG(tran_amount), 2) AS avg_amount
    FROM fact_transactions
    GROUP BY year
    ORDER BY year
""").fetchall()

print(f"  {'Year':<6} {'Transactions':>14} {'Revenue':>14} {'Avg Amount':>12}")
print("  " + "-"*50)
for row in result:
    print(f"  {row[0]:<6} {row[1]:>14,} ${row[2]:>13,} ${row[3]:>11}")

conn.close()

print(f"\nPhase 1 Complete!")
print(f"Database saved at: {DB_PATH}")