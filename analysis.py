import pandas as pd

# 讀取資料
df = pd.read_excel(r"C:\Users\user\Desktop\ecommerce-analysis\ecommerce_sales_data.xlsx")

# 瀏覽前 5 筆
print("前 5 筆資料：")
print(df.head())

# 檢查資料結構
print("\n資料結構：")
df.info() # 這裡不用 print()，info() 會自己輸出

# 檢查缺失值
print("\n缺失值統計：")
print(df.isnull().sum())

# 確認日期欄位型態，若不是 datetime 轉換
if df["Order Date"].dtype != "datetime64[ns]":
    df["Order Date"] = pd.to_datetime(df["Order Date"])
print("\n日期欄位型態：", df["Order Date"].dtype)



if df["Order Date"].dtype != "datetime64[ns]":
    df["Order Date"] = pd.to_datetime(df["Order Date"])

# === 1. 每月銷售趨勢 ===
# 建立 "Year-Month" 欄位
df["YearMonth"] = df["Order Date"].dt.to_period("M")
monthly_sales = df.groupby("YearMonth")["Sales"].sum()

print("\n=== 每月銷售趨勢 ===")
print(monthly_sales)

# === 2. Top 10 暢銷商品 ===
top_products = df.groupby("Product")["Sales"].sum().sort_values(ascending=False).head(10)

print("\n=== Top 10 暢銷商品 ===")
print(top_products)

import matplotlib.pyplot as plt


# === 折線圖：每月銷售趨勢 ===
plt.figure(figsize=(10,5))
monthly_sales.plot(kind="line", marker="o")
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Sales")
plt.grid(True)
plt.tight_layout()
plt.savefig("monthly_sales_trend.png")  # 存圖
plt.show()

# === 長條圖：Top 10 商品 ===
plt.figure(figsize=(10,5))
top_products.plot(kind="bar")
plt.title("Top 10 Products by Sales")
plt.xlabel("Product")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("top_10_products.png")  # 存圖
plt.show()



