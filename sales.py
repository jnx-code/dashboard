# Required pip packages:
# pip install pandas numpy matplotlib seaborn openpyxl

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage

# Read the sales data
sales_df = pd.read_csv('sales_data_sample.csv')

# Calculate KPIs
sales_df['Revenue'] = sales_df['SALES']
# Estimate cost as 70% of MSRP, profit as (Revenue - Estimated Cost)
sales_df['Estimated_Cost'] = sales_df['MSRP'] * 0.7 * sales_df['QUANTITYORDERED']
sales_df['Profit'] = sales_df['Revenue'] - sales_df['Estimated_Cost']

# Total Revenue and Profit
total_revenue = sales_df['Revenue'].sum()
total_profit = sales_df['Profit'].sum()

# Region-wise performance (by TERRITORY)
region_performance = sales_df.groupby('TERRITORY').agg({'Revenue': 'sum', 'Profit': 'sum'}).reset_index()

# Create Excel dashboard
excel_path = 'sales_dashboard.xlsx'
wb = Workbook()
ws_summary = wb.active
ws_summary.title = 'Summary'

# Write summary KPIs
ds = [
    ['Total Revenue', total_revenue],
    ['Total Profit', total_profit]
]
for row in ds:
    ws_summary.append(row)

ws_summary.append([])
ws_summary.append(['Region', 'Revenue', 'Profit'])
for _, row in region_performance.iterrows():
    ws_summary.append([row['TERRITORY'], row['Revenue'], row['Profit']])

# Save region performance as a chart
plt.figure(figsize=(8, 5))
sns.barplot(data=region_performance, x='TERRITORY', y='Revenue', color='skyblue')
plt.title('Revenue by Territory')
plt.ylabel('Revenue')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('region_revenue.png')
plt.close()

# Insert chart image into Excel
from openpyxl.drawing.image import Image as XLImage
img = XLImage('region_revenue.png')
ws_summary.add_image(img, 'E2')

# Save Excel file
wb.save(excel_path)

# Clean up chart image
os.remove('region_revenue.png')

print(f"Dashboard created: {excel_path}")
