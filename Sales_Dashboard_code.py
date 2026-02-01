import pandas as pd
import matplotlib.pyplot as plt

# إعداد الملف والشيتات 
file_path = r"C:\Users\frah\Desktop\Sales_Dashboard\Sales_Dashboard.xlsx"

# قراءة الجدول الرئيسي
df_main = pd.read_excel(file_path, sheet_name="SalesDataa", engine='openpyxl')

# قراءة الجدول الصغير
df_small = pd.read_excel(file_path, sheet_name="SalesDataa_SmallTable", engine='openpyxl')

print("Main table:")
print(df_main.head())
print("\nSmall table:")
print(df_small.head())

# تحويل التاريخ لصيغة datetime
df_main['Transaction Date'] = pd.to_datetime(df_main['Transaction Date'], dayfirst=True, errors='coerce')

# ---- 1. Sales by Region ----
sales_region = df_small.groupby('Region')['Net Amount'].sum().sort_values(ascending=False)
plt.figure(figsize=(8,5))
sales_region.plot(kind='bar', color='skyblue')
plt.title("Sales by Region")
plt.ylabel("Net Amount")
plt.tight_layout()
plt.savefig(r"C:\Users\frah\Desktop\Sales_Dashboard\Sales_by_Region.png")
plt.close()

# ---- 2. Sales by Customer Type ----
sales_customer_type = df_small.groupby('Customer Type')['Net Amount'].sum().sort_values(ascending=False)
plt.figure(figsize=(6,4))
sales_customer_type.plot(kind='bar', color='orange')
plt.title("Sales by Customer Type")
plt.ylabel("Net Amount")
plt.tight_layout()
plt.savefig(r"C:\Users\frah\Desktop\Sales_Dashboard\Sales_by_CustomerType.png")
plt.close()

# Sales by Sales Rep 
sales_rep = df_main.groupby('Sales Rep')['Net Amount'].sum().sort_values(ascending=False)
plt.figure(figsize=(8,5))
sales_rep.plot(kind='bar', color='green')
plt.title("Sales by Sales Rep")
plt.ylabel("Net Amount")
plt.tight_layout()
plt.savefig(r"C:\Users\frah\Desktop\Sales_Dashboard\Sales_by_SalesRep.png")
plt.close()

# ---- 4. Sales Trend (حسب التاريخ) ----
sales_trend = df_main.groupby('Transaction Date')['Net Amount'].sum()
plt.figure(figsize=(10,5))
sales_trend.plot(marker='o', linestyle='-')
plt.title("Sales Trend")
plt.ylabel("Net Amount")
plt.xlabel("Transaction Date")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r"C:\Users\frah\Desktop\Sales_Dashboard\Sales_Trend.png")
plt.close()

# ---- 5. Sales by Customer ----
sales_customer = df_small.set_index('Customer Name')['Net Amount'].sort_values(ascending=False)
plt.figure(figsize=(10,5))
sales_customer.plot(kind='bar', color='purple')
plt.title("Sales by Customer")
plt.ylabel("Net Amount")
plt.tight_layout()
plt.savefig(r"C:\Users\frah\Desktop\Sales_Dashboard\Sales_by_Customer.png")
plt.close()

print("All charts have been saved as PNG in your folder.")