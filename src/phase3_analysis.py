
# PHASE 3 — Data Analysis & Advanced Analytics
# Retail Chain Sales Analytics


import pandas as pd
import sqlite3
import os
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

# ── STEP 1: Paths & Setup ────────────────────────────────────
BASE_DIR = r"C:\Users\KIIT\OneDrive\Desktop\data_analytic\Sales Data Analysis and Reporting for a Retail Chain"
DATA_DIR = os.path.join(BASE_DIR, "data")
IMG_DIR  = os.path.join(BASE_DIR, "images")
DB_PATH  = os.path.join(DATA_DIR, "retail_analytics.db")

os.makedirs(IMG_DIR, exist_ok=True)

# Chart styling
sns.set_theme(style="darkgrid", font_scale=1.1)
plt.rcParams.update({
    "figure.facecolor" : "#0A0A0F",
    "axes.facecolor"   : "#12121A",
    "axes.labelcolor"  : "white",
    "xtick.color"      : "white",
    "ytick.color"      : "white",
    "text.color"       : "white",
    "grid.color"       : "#2a2a3a",
    "axes.spines.top"  : False,
    "axes.spines.right": False,
})
ORANGE = "#E85D04"
AMBER  = "#F48C06"
GREEN  = "#52B788"
BLUE   = "#48CAE4"
PALETTE = [ORANGE, AMBER, "#FFBA08", GREEN, BLUE,
           "#023E8A", "#D62828", "#A8DADC", "#E9C46A", "#457B9D"]

conn   = sqlite3.connect(DB_PATH)
cursor = conn.cursor()
df     = pd.read_sql_query(
    "SELECT * FROM fact_transactions_clean", conn
)
df['trans_date'] = pd.to_datetime(df['trans_date'])
print(f"Loaded {len(df):,} rows for analysis\n")



#  SECTION A — SQL EXPLORATION

print("=" * 55)
print("  SECTION A — SQL Exploration Queries")
print("=" * 55)

# A1: Total Revenue KPIs
print("\n[A1] Overall KPIs:")
r = cursor.execute("""
    SELECT
        COUNT(*)                        AS total_transactions,
        COUNT(DISTINCT customer_id)     AS unique_customers,
        SUM(tran_amount)                AS total_revenue,
        ROUND(AVG(tran_amount), 2)      AS avg_transaction,
        MAX(tran_amount)                AS max_transaction,
        MIN(tran_amount)                AS min_transaction
    FROM fact_transactions_clean
""").fetchone()
print(f"  Total Transactions : {r[0]:,}")
print(f"  Unique Customers   : {r[1]:,}")
print(f"  Total Revenue      : ${r[2]:,}")
print(f"  Avg Transaction    : ${r[3]}")
print(f"  Max Transaction    : ${r[4]}")
print(f"  Min Transaction    : ${r[5]}")

# A2: Revenue by Year
print("\n[A2] Revenue by Year:")
rows = cursor.execute("""
    SELECT year,
           COUNT(*)                AS transactions,
           SUM(tran_amount)        AS revenue,
           ROUND(AVG(tran_amount), 2) AS avg_amt,
           COUNT(DISTINCT customer_id) AS customers
    FROM fact_transactions_clean
    GROUP BY year ORDER BY year
""").fetchall()
print(f"  {'Year':<6} {'Txns':>8} {'Revenue':>12} {'Avg':>8} {'Customers':>10}")
print("  " + "-"*48)
for row in rows:
    print(f"  {row[0]:<6} {row[1]:>8,} ${row[2]:>11,} ${row[3]:>7} {row[4]:>10,}")

# A3: Revenue by Quarter
print("\n[A3] Revenue by Quarter:")
rows = cursor.execute("""
    SELECT year, quarter,
           COUNT(*)         AS transactions,
           SUM(tran_amount) AS revenue
    FROM fact_transactions_clean
    GROUP BY year, quarter
    ORDER BY year, quarter
""").fetchall()
for row in rows:
    print(f"  FY{row[0]} Q{row[1]}  |  {row[2]:>6,} txns  |  ${row[3]:>10,}")

# A4: Revenue by Day of Week
print("\n[A4] Revenue by Day of Week:")
rows = cursor.execute("""
    SELECT day_name, day_of_week,
           COUNT(*)          AS transactions,
           SUM(tran_amount)  AS revenue,
           ROUND(AVG(tran_amount),2) AS avg_amt
    FROM fact_transactions_clean
    GROUP BY day_name, day_of_week
    ORDER BY day_of_week
""").fetchall()
for row in rows:
    print(f"  {row[0]:<12} {row[2]:>7,} txns   ${row[3]:>10,}   avg ${row[4]}")

# A5: Customer Segment Revenue
print("\n[A5] Revenue by Customer Segment:")
rows = cursor.execute("""
    SELECT customer_segment,
           COUNT(DISTINCT customer_id) AS customers,
           SUM(tran_amount)            AS revenue,
           ROUND(AVG(tran_amount),2)   AS avg_txn
    FROM fact_transactions_clean
    GROUP BY customer_segment
    ORDER BY revenue DESC
""").fetchall()
for row in rows:
    print(f"  {row[0]:<12}  {row[1]:>5,} customers   ${row[2]:>10,}   avg ${row[3]}")

# A6: Response Rate Analysis
print("\n[A6] Response Rate Analysis:")
rows = cursor.execute("""
    SELECT responded,
           COUNT(DISTINCT customer_id) AS customers,
           SUM(tran_amount)            AS revenue,
           ROUND(AVG(tran_amount),2)   AS avg_txn
    FROM fact_transactions_clean
    GROUP BY responded
""").fetchall()
for row in rows:
    print(f"  Responded={row[0]:<4}  {row[1]:>5,} customers   ${row[2]:>10,}   avg ${row[3]}")

# A7: Top 10 Customers by Revenue
print("\n[A7] Top 10 Customers by Revenue:")
rows = cursor.execute("""
    SELECT customer_id,
           COUNT(*)         AS transactions,
           SUM(tran_amount) AS total_spent
    FROM fact_transactions_clean
    GROUP BY customer_id
    ORDER BY total_spent DESC
    LIMIT 10
""").fetchall()
for i, row in enumerate(rows, 1):
    print(f"  {i:>2}. {row[0]}   {row[1]:>4} txns   ${row[2]:>6,}")

# A8: Monthly Average Revenue
print("\n[A8] Average Revenue by Month (across all years):")
rows = cursor.execute("""
    SELECT month_num, month_name,
           ROUND(AVG(monthly_rev),2) AS avg_monthly_revenue
    FROM (
        SELECT month_num, month_name, year,
               SUM(tran_amount) AS monthly_rev
        FROM fact_transactions_clean
        GROUP BY month_num, month_name, year
    )
    GROUP BY month_num, month_name
    ORDER BY month_num
""").fetchall()
for row in rows:
    bar = "█" * int(row[2] / 5000)
    print(f"  {row[1]:<10} ${row[2]:>10,}  {bar}")



#  SECTION B — VISUALIZATIONS

print("\n" + "=" * 55)
print("  SECTION B — Generating Charts")
print("=" * 55)


# ── Chart 1: Monthly Revenue Trend ──────────────────────────
monthly = df.groupby(['year', 'month_num', 'month_name',
                      'year_month'])['tran_amount'].sum().reset_index()
monthly = monthly.sort_values(['year', 'month_num'])

fig, ax = plt.subplots(figsize=(14, 5))
ax.plot(monthly['year_month'], monthly['tran_amount'],
        color=ORANGE, linewidth=2.5, marker='o', markersize=5)
ax.fill_between(monthly['year_month'], monthly['tran_amount'],
                alpha=0.15, color=ORANGE)
ax.set_title("Monthly Revenue Trend (2011–2015)",
             fontsize=14, fontweight='bold', color='white', pad=14)
ax.set_xlabel("Month")
ax.set_ylabel("Revenue ($)")
ax.yaxis.set_major_formatter(
    mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}K"))
plt.xticks(rotation=45, ha='right', fontsize=8)
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "01_monthly_revenue.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 1: Monthly Revenue saved")


# ── Chart 2: Revenue by Year (Bar) ──────────────────────────
yearly = df.groupby('year')['tran_amount'].sum().reset_index()

fig, ax = plt.subplots(figsize=(8, 5))
bars = ax.bar(yearly['year'].astype(str), yearly['tran_amount'],
              color=PALETTE[:len(yearly)], width=0.5,
              edgecolor='white', linewidth=0.5)
ax.bar_label(bars,
             labels=[f"${v/1000:.0f}K" for v in yearly['tran_amount']],
             padding=5, color='white', fontsize=10)
ax.set_title("Total Revenue by Year",
             fontsize=14, fontweight='bold', color='white', pad=14)
ax.set_xlabel("Year")
ax.set_ylabel("Revenue ($)")
ax.yaxis.set_major_formatter(
    mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}K"))
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "02_yearly_revenue.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 2: Yearly Revenue saved")


# ── Chart 3: Revenue by Day of Week ─────────────────────────
dow = df.groupby(['day_of_week', 'day_name'])['tran_amount']\
        .sum().reset_index().sort_values('day_of_week')

fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar(dow['day_name'], dow['tran_amount'],
              color=[GREEN if i < 5 else AMBER
                     for i in range(len(dow))],
              edgecolor='white', linewidth=0.5)
ax.bar_label(bars,
             labels=[f"${v/1000:.0f}K" for v in dow['tran_amount']],
             padding=4, color='white', fontsize=9)
ax.set_title("Revenue by Day of Week",
             fontsize=14, fontweight='bold', color='white', pad=14)
ax.set_xlabel("Day")
ax.set_ylabel("Revenue ($)")
ax.yaxis.set_major_formatter(
    mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}K"))
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "03_revenue_by_day.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 3: Revenue by Day saved")


# ── Chart 4: Transaction Amount Distribution ─────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 5))

axes[0].hist(df['tran_amount'], bins=30,
             color=ORANGE, edgecolor='white', linewidth=0.5, alpha=0.85)
axes[0].set_title("Transaction Amount Distribution",
                  color='white', fontweight='bold')
axes[0].set_xlabel("Amount ($)")
axes[0].set_ylabel("Frequency")
axes[0].set_facecolor("#12121A")

sns.boxplot(y=df['tran_amount'], ax=axes[1], color=AMBER)
axes[1].set_title("Boxplot — Transaction Amount",
                  color='white', fontweight='bold')
axes[1].set_ylabel("Amount ($)")
axes[1].set_facecolor("#12121A")

fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "04_amount_distribution.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 4: Amount Distribution saved")


# ── Chart 5: Customer Segment Revenue ───────────────────────
seg = df.groupby('customer_segment')['tran_amount']\
        .sum().reset_index().sort_values('tran_amount', ascending=False)

fig, ax = plt.subplots(figsize=(8, 5))
wedges, texts, autotexts = ax.pie(
    seg['tran_amount'],
    labels=seg['customer_segment'],
    colors=PALETTE[:len(seg)],
    autopct='%1.1f%%',
    startangle=140,
    wedgeprops=dict(edgecolor='#0A0A0F', linewidth=2)
)
for t in texts + autotexts:
    t.set_color('white')
ax.set_title("Revenue Share by Customer Segment",
             fontsize=13, fontweight='bold', color='white', pad=14)
fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "05_customer_segments.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 5: Customer Segments saved")


# ── Chart 6: Monthly Avg Revenue Heatmap ────────────────────
pivot = df.groupby(['year', 'month_num'])['tran_amount']\
          .sum().reset_index()
pivot = pivot.pivot(index='year', columns='month_num',
                    values='tran_amount').fillna(0)
pivot.columns = ['Jan','Feb','Mar','Apr','May','Jun',
                 'Jul','Aug','Sep','Oct','Nov','Dec'][:len(pivot.columns)]

fig, ax = plt.subplots(figsize=(14, 4))
sns.heatmap(pivot / 1000, annot=True, fmt='.0f',
            cmap='YlOrRd', linewidths=0.5,
            linecolor='#0A0A0F', ax=ax,
            cbar_kws={'label': 'Revenue ($K)'})
ax.set_title("Revenue Heatmap — Year × Month ($K)",
             fontsize=13, fontweight='bold', color='white', pad=14)
ax.set_ylabel("Year")
ax.set_xlabel("Month")
fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "06_heatmap_year_month.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 6: Heatmap saved")


# ── Chart 7: Response vs No Response Revenue ─────────────────
resp_grp = df.groupby('responded')['tran_amount']\
             .agg(['sum', 'mean', 'count']).reset_index()

fig, axes = plt.subplots(1, 2, figsize=(12, 5))
colors = [ORANGE, GREEN]

axes[0].bar(resp_grp['responded'], resp_grp['sum'] / 1000,
            color=colors, edgecolor='white', linewidth=0.5)
axes[0].set_title("Total Revenue: Responded vs Not",
                  color='white', fontweight='bold')
axes[0].set_ylabel("Revenue ($K)")
axes[0].set_facecolor("#12121A")
for i, v in enumerate(resp_grp['sum']):
    axes[0].text(i, v/1000 + 20, f"${v/1000:.0f}K",
                 ha='center', color='white', fontsize=10)

axes[1].bar(resp_grp['responded'], resp_grp['mean'],
            color=colors, edgecolor='white', linewidth=0.5)
axes[1].set_title("Avg Transaction: Responded vs Not",
                  color='white', fontweight='bold')
axes[1].set_ylabel("Avg Amount ($)")
axes[1].set_facecolor("#12121A")
for i, v in enumerate(resp_grp['mean']):
    axes[1].text(i, v + 0.3, f"${v:.2f}",
                 ha='center', color='white', fontsize=10)

fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "07_response_analysis.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 7: Response Analysis saved")


#  SECTION C — ADVANCED ANALYTICS

print("\n" + "=" * 55)
print("  SECTION C — Advanced Analytics")
print("=" * 55)


# ── C1: RFM Analysis ─────────────────────────────────────────
print("\n[C1] RFM Customer Segmentation...")

snapshot = df['trans_date'].max() + pd.Timedelta(days=1)
rfm = df.groupby('customer_id').agg(
    recency   = ('trans_date',   lambda x: (snapshot - x.max()).days),
    frequency = ('trans_date',   'count'),
    monetary  = ('tran_amount',  'sum')
).reset_index()

rfm['R'] = pd.qcut(rfm['recency'],
                   5, labels=[5,4,3,2,1]).astype(int)
rfm['F'] = pd.qcut(rfm['frequency'].rank(method='first'),
                   5, labels=[1,2,3,4,5]).astype(int)
rfm['M'] = pd.qcut(rfm['monetary'].rank(method='first'),
                   5, labels=[1,2,3,4,5]).astype(int)

def rfm_segment(row):
    r, f, m = row['R'], row['F'], row['M']
    if r >= 4 and f >= 4 and m >= 4: return 'Champions'
    if r >= 3 and f >= 3:            return 'Loyal Customers'
    if r >= 4 and f <= 2:            return 'Recent Customers'
    if r >= 3:                       return 'Promising'
    if r <= 2 and f >= 3:            return 'At Risk'
    if r <= 1 and f <= 2:            return 'Lost'
    return 'Need Attention'

rfm['segment'] = rfm.apply(rfm_segment, axis=1)

print("\n  RFM Segment Summary:")
seg_summary = rfm.groupby('segment').agg(
    customers = ('customer_id', 'count'),
    avg_recency   = ('recency',   'mean'),
    avg_frequency = ('frequency', 'mean'),
    avg_monetary  = ('monetary',  'mean')
).round(1)
print(seg_summary.to_string())

# RFM Chart
seg_counts = rfm['segment'].value_counts()
fig, axes = plt.subplots(1, 2, figsize=(14, 5))

axes[0].barh(seg_counts.index, seg_counts.values,
             color=PALETTE[:len(seg_counts)])
axes[0].set_title("Customer Count by RFM Segment",
                  color='white', fontweight='bold')
axes[0].set_xlabel("Number of Customers")
axes[0].set_facecolor("#12121A")

seg_rev = rfm.groupby('segment')['monetary'].sum().sort_values()
axes[1].barh(seg_rev.index, seg_rev.values / 1000,
             color=PALETTE[:len(seg_rev)])
axes[1].set_title("Revenue by RFM Segment ($K)",
                  color='white', fontweight='bold')
axes[1].set_xlabel("Total Revenue ($K)")
axes[1].set_facecolor("#12121A")

fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "08_rfm_segments.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 8: RFM Segments saved")


# ── C2: Cohort Analysis ──────────────────────────────────────
print("\n[C2] Cohort Retention Analysis...")

first_year = df.groupby('customer_id')['year']\
               .min().rename('cohort_year')
df2 = df.join(first_year, on='customer_id')

cohort = df2.groupby(['cohort_year', 'year'])\
            ['customer_id'].nunique().reset_index()
cohort_pivot = cohort.pivot(
    index='cohort_year', columns='year',
    values='customer_id'
).fillna(0)
cohort_size    = cohort_pivot.iloc[:, 0]
retention      = (cohort_pivot.divide(cohort_size, axis=0) * 100).round(1)

print("\n  Cohort Retention Matrix (%):")
print(retention.to_string())

fig, ax = plt.subplots(figsize=(10, 5))
sns.heatmap(retention, annot=True, fmt='.0f',
            cmap='YlOrRd', linewidths=0.5,
            linecolor='#0A0A0F', ax=ax,
            cbar_kws={'label': 'Retention %'})
ax.set_title("Customer Cohort Retention by Year (%)",
             fontsize=13, fontweight='bold', color='white', pad=14)
ax.set_xlabel("Year")
ax.set_ylabel("Cohort Year")
fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "09_cohort_retention.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 9: Cohort Retention saved")


# ── C3: Time Series — Rolling Average ───────────────────────
print("\n[C3] Time Series Analysis...")

ts = df.groupby('trans_date')['tran_amount']\
       .sum().reset_index().sort_values('trans_date')
ts['rolling_30'] = ts['tran_amount'].rolling(30).mean()
ts['rolling_90'] = ts['tran_amount'].rolling(90).mean()

fig, ax = plt.subplots(figsize=(14, 5))
ax.plot(ts['trans_date'], ts['tran_amount'],
        color=ORANGE, alpha=0.3, linewidth=1, label='Daily Revenue')
ax.plot(ts['trans_date'], ts['rolling_30'],
        color=AMBER, linewidth=2, label='30-Day Rolling Avg')
ax.plot(ts['trans_date'], ts['rolling_90'],
        color=GREEN, linewidth=2.5, label='90-Day Rolling Avg')
ax.set_title("Daily Revenue with Rolling Averages",
             fontsize=14, fontweight='bold', color='white', pad=14)
ax.set_xlabel("Date")
ax.set_ylabel("Revenue ($)")
ax.yaxis.set_major_formatter(
    mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}K"))
ax.legend(facecolor='#12121A', labelcolor='white')
fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "10_time_series.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 10: Time Series saved")


# ── C4: Customer LTV Distribution ───────────────────────────
print("\n[C4] Customer Lifetime Value Distribution...")

ltv = df.groupby('customer_id')['tran_amount'].sum().reset_index()
ltv.columns = ['customer_id', 'ltv']

fig, ax = plt.subplots(figsize=(10, 5))
ax.hist(ltv['ltv'], bins=40,
        color=BLUE, edgecolor='white', linewidth=0.5, alpha=0.85)
ax.axvline(ltv['ltv'].mean(), color=ORANGE,
           linestyle='--', linewidth=2,
           label=f"Mean LTV: ${ltv['ltv'].mean():,.0f}")
ax.axvline(ltv['ltv'].median(), color=GREEN,
           linestyle='--', linewidth=2,
           label=f"Median LTV: ${ltv['ltv'].median():,.0f}")
ax.set_title("Customer Lifetime Value Distribution",
             fontsize=13, fontweight='bold', color='white', pad=14)
ax.set_xlabel("Lifetime Value ($)")
ax.set_ylabel("Number of Customers")
ax.legend(facecolor='#12121A', labelcolor='white')
fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "11_ltv_distribution.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 11: LTV Distribution saved")


# ── C5: Quarterly Growth Rate ────────────────────────────────
print("\n[C5] Quarterly Growth Rate...")

qtr = df.groupby(['year', 'quarter'])['tran_amount']\
        .sum().reset_index()
qtr['label']  = qtr['year'].astype(str) + '-Q' + qtr['quarter'].astype(str)
qtr['growth'] = qtr['tran_amount'].pct_change() * 100

fig, ax1 = plt.subplots(figsize=(12, 5))
ax2 = ax1.twinx()

ax1.bar(qtr['label'], qtr['tran_amount'] / 1000,
        color=ORANGE, alpha=0.7, label='Revenue ($K)')
ax2.plot(qtr['label'], qtr['growth'],
         color=GREEN, marker='o', linewidth=2,
         markersize=6, label='QoQ Growth %')
ax2.axhline(0, color='white', linestyle='--', alpha=0.3)

ax1.set_title("Quarterly Revenue & Growth Rate",
              fontsize=13, fontweight='bold', color='white', pad=14)
ax1.set_xlabel("Quarter")
ax1.set_ylabel("Revenue ($K)", color=ORANGE)
ax2.set_ylabel("QoQ Growth (%)", color=GREEN)
ax1.set_facecolor("#12121A")

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2,
           facecolor='#12121A', labelcolor='white')

plt.xticks(rotation=45, ha='right')
fig.patch.set_facecolor("#0A0A0F")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "12_quarterly_growth.png"),
            dpi=150, bbox_inches='tight')
plt.show()
print("  ✔ Chart 12: Quarterly Growth saved")



#  FINAL SUMMARY

print("\n" + "=" * 55)
print("  PHASE 3 COMPLETE — Summary")
print("=" * 55)
print(f"\n  Charts saved to: {IMG_DIR}")
charts = [
    "01_monthly_revenue.png",
    "02_yearly_revenue.png",
    "03_revenue_by_day.png",
    "04_amount_distribution.png",
    "05_customer_segments.png",
    "06_heatmap_year_month.png",
    "07_response_analysis.png",
    "08_rfm_segments.png",
    "09_cohort_retention.png",
    "10_time_series.png",
    "11_ltv_distribution.png",
    "12_quarterly_growth.png",
]
for c in charts:
    print(f"  ✔ {c}")

rfm.to_sql('rfm_segments', conn, if_exists='replace', index=False)
conn.commit()
conn.close()

print("\nPhase 3 Complete!")
print("   RFM results saved to database: rfm_segments table")
print("   Ready for Phase 4 — Reporting & Dashboard")