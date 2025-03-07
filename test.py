import pandas as pd

# 讀取兩個 Excel 檔案
file_path_a = "./a.xlsx"  # 收款金額表
file_path_b = "./b.xlsx"  # 銷售金額表
df_a = pd.read_excel(file_path_a)
df_b = pd.read_excel(file_path_b)

# 清理欄位名稱（移除前後空格）
df_a.columns = df_a.columns.str.strip()
df_b.columns = df_b.columns.str.strip()

# 找到合計行的索引，並只保留合計行之前的資料
total_row_index = df_a[df_a['部門'] == '合計'].index[0]
df_a = df_a.iloc[:total_row_index]  # 只取合計行之前的資料

# 先填充空白的營業日期（向下填充）
df_a['營業日期'] = df_a['營業日期'].fillna(method='ffill')
df_b['營業日期'] = df_b['營業日期'].fillna(method='ffill')

# 確保兩個表格的 '營業日期' 欄位為 datetime 格式
df_a['營業日期'] = pd.to_datetime(df_a['營業日期'])
df_b['營業日期'] = pd.to_datetime(df_b['營業日期'])

# 只保留日期（去除時間部分）
df_a['日期'] = df_a['營業日期'].dt.date
df_b['日期'] = df_b['營業日期'].dt.date

# 診斷特定日期的資料
print("\n=== A表格 1/25 的詳細資料 ===")
target_date = pd.to_datetime('2025-1-25').date()
print(df_a[df_a['日期'] == target_date][['營業日期', '部門', '收款金額']])
print("\n收款金額總和:", df_a[df_a['日期'] == target_date]['收款金額'].sum())

print("\n=== B表格 1/25 的詳細資料 ===")
print(df_b[df_b['日期'] == target_date][['營業日期', '銷售金額']])
print("\n銷售金額總和:", df_b[df_b['日期'] == target_date]['銷售金額'].sum())

# 分別計算兩個表格的每日總金額
daily_total_a = df_a.groupby('日期')['收款金額'].sum().reset_index()
daily_total_b = df_b.groupby('日期')['銷售金額'].sum().reset_index()

# 合併兩個結果，使用 outer join 確保不會遺漏任一表格的日期
merged_totals = pd.merge(
    daily_total_a, 
    daily_total_b, 
    on='日期', 
    how='outer'
)

# 計算差異並標記
merged_totals['金額是否相符'] = (merged_totals['收款金額'] == merged_totals['銷售金額'])
merged_totals['差額'] = merged_totals['收款金額'] - merged_totals['銷售金額']

# 顯示結果
print(merged_totals)

# 存成新的 Excel 檔案
merged_totals.to_excel("comparison_result.xlsx", index=False)