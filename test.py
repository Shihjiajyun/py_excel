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
df_a['營業日期'] = df_a['營業日期'].ffill()
df_b['營業日期'] = df_b['營業日期'].ffill()

# 填充空白的桌號/台位（向下填充）
df_a['桌號'] = df_a['桌號'].ffill()
df_b['臺位'] = df_b['臺位'].ffill()

# 確保兩個表格的 '營業日期' 欄位為 datetime 格式
df_a['營業日期'] = pd.to_datetime(df_a['營業日期'])
df_b['營業日期'] = pd.to_datetime(df_b['營業日期'])

# 只保留日期（去除時間部分）
df_a['日期'] = df_a['營業日期'].dt.date
df_b['日期'] = df_b['營業日期'].dt.date

# 診斷特定日期的資料
print("\n=== 指定日期的詳細資料比對 ===")
target_date = pd.to_datetime('2025-1-25').date()

# 分別計算兩個表格的每日總金額
daily_total_a = df_a.groupby('日期')['收款金額'].sum().reset_index()
daily_total_b = df_b.groupby('日期')['銷售金額'].sum().reset_index()

# 合併每日總金額比對
merged_totals = pd.merge(
    daily_total_a, 
    daily_total_b, 
    on='日期', 
    how='outer'
).fillna(0)

# 計算總金額差異
merged_totals['金額是否相符'] = (merged_totals['收款金額'] == merged_totals['銷售金額'])
merged_totals['差額'] = merged_totals['收款金額'] - merged_totals['銷售金額']

# 處理位置明細比對
# 統一位置編號的格式
def standardize_position(pos):
    pos = str(pos).strip()
    # 移除所有空格
    pos = pos.replace(' ', '')
    # 統一全形/半形
    pos = pos.replace('　', '')
    return pos

# 先確保位置編號的格式一致，並填充空白值
df_a['位置編號'] = df_a['桌號'].astype(str).apply(standardize_position)
df_b['位置編號'] = df_b['臺位'].astype(str).apply(standardize_position)

# 分別計算每個位置的總金額
detail_a = df_a.groupby(['日期', '位置編號'])['收款金額'].sum().reset_index()
detail_b = df_b.groupby(['日期', '位置編號'])['銷售金額'].sum().reset_index()

# 合併同一天同一位置的資料
position_comparison = pd.merge(
    detail_a,
    detail_b,
    on=['日期', '位置編號'],
    how='outer'
).fillna(0)

# 計算每個位置的差異
position_comparison['金額是否相符'] = (position_comparison['收款金額'] == position_comparison['銷售金額'])
position_comparison['差額'] = position_comparison['收款金額'] - position_comparison['銷售金額']

# 排序結果
merged_totals = merged_totals.sort_values('日期')
position_comparison = position_comparison.sort_values(['日期', '位置編號'])

# 在準備詳細資料之前，先清理桌號格式
def clean_table_number(table_num):
    if pd.isna(table_num):
        return table_num
    return str(table_num).strip().replace(' ', '').replace('　', '')

# 清理A表和B表的桌號
df_a['桌號'] = df_a['桌號'].apply(clean_table_number)
df_b['臺位'] = df_b['臺位'].apply(clean_table_number)

# 準備A表的詳細資料
a_details = df_a.groupby(['營業日期', '桌號']).agg({
    '付款方式': lambda x: '、'.join(x.dropna().astype(str).unique()),
    '付款金額': 'sum',
    '發票號碼': lambda x: '、'.join(x.dropna().astype(str).unique()),
    '收款金額': 'sum',
    '未稅': 'sum',
    '稅額': 'sum'
}).reset_index()

# 準備B表的詳細資料
b_details = df_b.groupby(['營業日期', '臺位']).agg({
    '餐飲類別': lambda x: '、'.join(x.dropna().astype(str).unique()),
    '品名': lambda x: '、'.join(x.dropna().astype(str).unique()),
    '單位': lambda x: '、'.join(x.dropna().astype(str).unique()),
    '銷售數量': 'sum',
    '銷售金額': 'sum'
}).reset_index()

# 將B表的'臺位'改名為'桌號'以便合併
b_details = b_details.rename(columns={'臺位': '桌號'})

# 合併A表和B表的詳細資料
detailed_comparison = pd.merge(
    a_details,
    b_details,
    on=['營業日期', '桌號'],
    how='outer'
)

# 處理空值
# 數值欄位填0，文字欄位填空字串
numeric_columns = ['付款金額', '收款金額', '未稅', '稅額', '銷售數量', '銷售金額']
text_columns = ['付款方式', '發票號碼', '餐飲類別', '品名', '單位']

for col in numeric_columns:
    detailed_comparison[col] = detailed_comparison[col].fillna(0)
for col in text_columns:
    detailed_comparison[col] = detailed_comparison[col].fillna('')

# 計算金額差異
detailed_comparison['金額是否相符'] = (detailed_comparison['收款金額'] == detailed_comparison['銷售金額'])
detailed_comparison['差額'] = detailed_comparison['收款金額'] - detailed_comparison['銷售金額']

# 排序結果
detailed_comparison = detailed_comparison.sort_values(['營業日期', '桌號'])

# 將結果存入Excel的不同分頁
with pd.ExcelWriter("comparison_result.xlsx") as writer:
    # 第一個分頁：只有日期總額對照
    merged_totals.to_excel(writer, sheet_name='日期總額比對', index=False)
    
    # 第二個分頁：同一天同一位置的詳細比對
    position_comparison.to_excel(writer, sheet_name='位置明細比對', index=False)
    
    # 第三個分頁：詳細資訊比對
    detailed_comparison.to_excel(writer, sheet_name='詳細資訊比對', index=False)

# 顯示結果
print("\n=== 日期總額比對 ===")
print(merged_totals)
print("\n=== 位置明細比對 ===")
print(position_comparison)
print("\n=== 詳細資訊比對 ===")
print(detailed_comparison)