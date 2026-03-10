#!/usr/bin/env python3
"""
测试数据格式保持功能
"""

import pandas as pd
from excel_validator import process_excel_with_validation

# 创建测试数据
test_data = {
    '订单号': ['001', '002', '003', '001', '002', '003'],
    '产品代码': ['P01', 'P02', 'P03', 'P01', 'P02', 'P03'],
    '部门': ['A', 'A', 'B', 'B', 'A', 'B'],
    '月份': ['2024-01', '2024-01', '2024-01', '2024-02', '2024-02', '2024-02'],
    '计划数量': [100, 200, 150, 120, 180, 160],
    '实际数量': [100, 200, 150, 120, 180, 160]
}

# 保存测试数据到Excel
test_df = pd.DataFrame(test_data)
test_df.to_excel('test_data.xlsx', index=False)

print("测试数据创建完成:")
print(test_df)
print("\n订单号的数据类型:", test_df['订单号'].dtype)
print("产品代码的数据类型:", test_df['产品代码'].dtype)

# 测试1：不使用string_columns参数
print("\n=== 测试1：不使用string_columns参数 ===")
result1 = process_excel_with_validation(
    input_file='test_data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门', '月份'],
    compare_columns=['计划数量', '实际数量'],
    output_columns=['部门', '月份', '订单号', '产品代码'],
    output_file='test_result_without_string.xlsx'
)

print("结果1:")
print(result1)
print("订单号的数据类型:", result1['订单号'].dtype if '订单号' in result1.columns else "不存在")

# 测试2：使用string_columns参数
print("\n=== 测试2：使用string_columns参数 ===")
result2 = process_excel_with_validation(
    input_file='test_data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门', '月份'],
    compare_columns=['计划数量', '实际数量'],
    output_columns=['部门', '月份', '订单号', '产品代码'],
    output_file='test_result_with_string.xlsx',
    string_columns=['订单号', '产品代码']
)

print("结果2:")
print(result2)
print("订单号的数据类型:", result2['订单号'].dtype if '订单号' in result2.columns else "不存在")