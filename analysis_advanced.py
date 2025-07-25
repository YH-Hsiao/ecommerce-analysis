import pandas as pd
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False
import os

# === 建立輸出資料夾 ===
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

# === 讀取資料 ===
df = pd.read_excel(r"C:\Users\user\Desktop\ecommerce-analysis\ecommerce_sales_data.xlsx")

# === 日期格式處理 ===
if df["Order Date"].dtype != "datetime64[ns]":
    df["Order Date"] = pd.to_datetime(df["Order Date"])

# === Year-Month 欄位 ===
df["YearMonth"] = df["Order Date"].dt.to_period("M")

# =========================
# 1. 每月銷售趨勢
# =========================
monthly_sales = df.groupby("YearMonth")["Sales"].sum()
plt.figure(figsize=(10, 5))
monthly_sales.plot(marker="o")
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Sales")
plt.grid(True)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "Monthly_Sales_Trend.png"))
plt.close()

# =========================
# 2. Top 10 暢銷商品
# =========================
top_products = df.groupby("Product")["Sales"].sum().sort_values(ascending=False).head(10)
plt.figure(figsize=(10, 5))
top_products.plot(kind="bar")
plt.title("Top 10 Products by Sales")
plt.xlabel("Product")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "Top_10_Products.png"))
plt.close()

# =========================
# 3. 產品類別銷售佔比
# =========================
category_sales = df.groupby("Category")["Sales"].sum()
plt.figure(figsize=(6, 6))
category_sales.plot(kind="pie", autopct='%1.1f%%', startangle=90)
plt.title("Sales by Category")
plt.ylabel("")
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "Category_Sales_Pie.png"))
plt.close()

# =========================
# 4. 週銷售趨勢（找出哪天最好）
# =========================
df['Weekday'] = df['Order Date'].dt.day_name()
weekday_sales = df.groupby('Weekday')['Sales'].sum()
weekday_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
weekday_sales = weekday_sales.reindex(weekday_order)

plt.figure(figsize=(8, 5))
weekday_sales.plot(kind='bar')
plt.title('Sales by Weekday')
plt.ylabel('Sales')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "Sales_by_Weekday.png"))
plt.close()

# =========================
# 5. 移動平均（7 天）趨勢
# =========================
daily_sales = df.resample('D', on='Order Date')['Sales'].sum()
rolling_sales = daily_sales.rolling(window=7).mean()

plt.figure(figsize=(10, 5))
plt.plot(daily_sales.index, daily_sales.values, label='Daily Sales')
plt.plot(rolling_sales.index, rolling_sales.values, label='7-Day Moving Average', linewidth=2)
plt.legend()
plt.title('Daily Sales with 7-Day Moving Average')
plt.ylabel('Sales')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "Daily_Sales_Moving_Average.png"))
plt.close()

# === 輸出 CSV ===
monthly_sales.to_csv(os.path.join(output_dir, "monthly_sales.csv"))
top_products.to_csv(os.path.join(output_dir, "top_10_products.csv"))
category_sales.to_csv(os.path.join(output_dir, "category_sales.csv"))
weekday_sales.to_csv(os.path.join(output_dir, "weekday_sales.csv"))
print("CSV 輸出完成！")
print("進階分析完成！所有圖表已輸出到 output/ 資料夾。")

# =========================
# 6. 產品層級的 RFM 分群分析（修正版）
# =========================

# 設定今天時間為最後一筆訂單日 + 1 天
today_date = df["Order Date"].max() + pd.Timedelta(days=1)

# 對每個產品做 RFM 分析（避免欄位名稱重複）
rfm_product = df.groupby("Product").agg(
    Recency=("Order Date", lambda x: (today_date - x.max()).days),
    Frequency=("Order Date", "count"),
    Monetary=("Sales", "sum")
)

# 打分數（1~5 分）
rfm_product["R_Score"] = pd.qcut(rfm_product["Recency"], 5, labels=[5,4,3,2,1])
rfm_product["F_Score"] = pd.qcut(rfm_product["Frequency"].rank(method="first"), 5, labels=[1,2,3,4,5])
rfm_product["M_Score"] = pd.qcut(rfm_product["Monetary"], 5, labels=[1,2,3,4,5])

# 總分
rfm_product["RFM_Score"] = (
    rfm_product["R_Score"].astype(int) +
    rfm_product["F_Score"].astype(int) +
    rfm_product["M_Score"].astype(int)
)

# 分群函數
def segment_product(score):
    if score >= 12:
        return "明星商品"
    elif score >= 9:
        return "潛力商品"
    elif score >= 6:
        return "普通商品"
    else:
        return "需淘汰"

rfm_product["Segment"] = rfm_product["RFM_Score"].apply(segment_product)

# 輸出結果
rfm_product.to_csv(os.path.join(output_dir, "rfm_products.csv"))
print("產品 RFM 分析完成！結果已輸出為 rfm_products.csv")

# =========================
# 7. RFM 結果視覺化
# =========================

import seaborn as sns

# 讀取剛剛的 RFM 結果
rfm_products = pd.read_csv(os.path.join(output_dir, "rfm_products.csv"), index_col=0)

# === 1. Frequency vs Monetary 散點圖（用顏色標分群） ===
plt.figure(figsize=(8, 6))
sns.scatterplot(
    data=rfm_products,
    x="Frequency",
    y="Monetary",
    hue="Segment",
    palette="Set2",
    s=80
)
plt.title("產品分群：Frequency vs Monetary")
plt.xlabel("購買次數 (Frequency)")
plt.ylabel("銷售額 (Monetary)")
plt.legend(title="分群")
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "RFM_Scatter.png"))
plt.close()

# === 2. 各分群商品數量長條圖 ===
plt.figure(figsize=(6, 4))
segment_counts = rfm_products["Segment"].value_counts()
sns.barplot(x=segment_counts.index, y=segment_counts.values)
plt.title("各分群商品數量")
plt.ylabel("數量")
plt.xlabel("分群")
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "RFM_Segment_Count.png"))
plt.close()

print("Day8 完成！已輸出 RFM_Scatter.png 與 RFM_Segment_Count.png")