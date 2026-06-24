import pandas as pd
import sqlite3
import os

BASE_DIR = r"C:\Users\KIIT\OneDrive\Desktop\data_analytic\Sales Data Analysis and Reporting for a Retail Chain"
DB_PATH  = os.path.join(BASE_DIR, "data", "retail_analytics.db")
RPT_DIR  = os.path.join(BASE_DIR, "reports")
os.makedirs(RPT_DIR, exist_ok=True)

conn = sqlite3.connect(DB_PATH)

# Export main cleaned data
df = pd.read_sql_query("SELECT * FROM fact_transactions_clean", conn)
df.to_excel(os.path.join(RPT_DIR, "cleaned_data.xlsx"), index=False)

# Export RFM segments
rfm = pd.read_sql_query("SELECT * FROM rfm_segments", conn)
rfm.to_excel(os.path.join(RPT_DIR, "rfm_data.xlsx"), index=False)

conn.close()
