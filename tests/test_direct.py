#!/usr/bin/env python3
"""
直接测试：创建一个能清晰展示 string_columns 效果的用例
"""

import pandas as pd
from excel_validator import process_excel_with_validation
import os

# 创建测试数据 - 确保数据类型明确
print("=== 直接测试 string_columns 效果 ===\n")

# 创建数据，确保某些列是数字格式，然后测试 string_columns 是否能将它们转为字符串
test_data = {
    # 这些列应该是数字格式
    '部门ID': [1, 1, 2, 2, 3, 3],  # int64
    '客户ID': [101, 102, 103, 104, 105, 106],  # int64
    '订单金额': [1000, 2000, 1500, 3000, 2500, 1800],  # int64
    '实际支付': [1000, 2000, 1500, 3100, 2500, 1800],  # int64，第4行有异常
    # 这些列已经是字符串格式
    '订单号': ['ORD-001', 'ORD-002', 'ORD-003', 'ORD-004', 'ORD-005', 'ORD-006'],
    '产品名称': ['产品A', '产品B', '产品C', '产品A', '产品B', '产品C']
}

# 确保列类型是我们想要的
test_df = pd.DataFrame(test_data)
test_df['部门ID'] = test_df['部门ID'].astype(int)
test_df['客户ID'] = test_df['客户ID'].astype(int)

test_df.to_excel('direct_test.xlsx', index=False)

print("1. 原始数据预览：")
print("数据类型：")
for col in test_df.columns:
    print(f"  {col}: {test_df[col].dtype}")

print("\n数据内容：")
print(test_df)

print("\n2. 分组逻辑：")
print("- 分组列：部门ID（数字）")
print("- 比较列：订单金额 vs 实际支付")
print("- 预期：部门ID=2的组应该异常（因为订单104有异常）")

# 测试：不使用 string_columns
print("\n3. 测试：不使用 string_columns")
result_no_string = process_excel_with_validation(
    input_file='direct_test.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门ID'],
    compare_columns=['订单金额', '实际支付'],
    output_columns=['部门ID', '验证状态', '总行数'],
    output_file='result_no_string.xlsx'
)

print("结果（不使用 string_columns）：")
print("数据类型：")
for col in result_no_string.columns:
    print(f"  {col}: {result_no_string[col].dtype}")
print("\n数据内容：")
print(result_no_string)

# 测试：使用 string_columns
print("\n4. 测试：使用 string_columns 将部门ID和客户ID转为字符串")
result_with_string = process_excel_with_validation(
    input_file='direct_test.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门ID'],
    compare_columns=['订单金额', '实际支付'],
    output_columns=['部门ID', '验证状态', '总行数'],
    output_file='result_with_string.xlsx',
    string_columns=['部门ID', '客户ID']
)

print("结果（使用 string_columns=['部门ID', '客户ID']）：")
print("数据类型：")
for col in result_with_string.columns:
    print(f"  {col}: {result_with_string[col].dtype}")
print("\n数据内容：")
print(result_with_string)

# 检查异常详情中的 string_columns 效果
print("\n5. 检查异常详情中的数据格式")
try:
    abnormal_detail = pd.read_excel('result_with_string.xlsx', sheet_name='异常详情')
    print("异常详情：")
    print("数据类型：")
    for col in abnormal_detail.columns:
        print(f"  {col}: {abnormal_detail[col].dtype}")
    print("\n数据内容：")
    print(abnormal_detail)

    # 检查订单号是否保持了字符串格式
    if '订单号' in abnormal_detail.columns:
        print(f"\n订单号类型检查：{abnormal_detail['订单号'].dtype}")
        print("订单号内容：", abnormal_detail['订单号'].tolist())
    else:
        print("\n❌ 订单号没有在异常详情中")

except Exception as e:
    print(f"读取异常详情失败：{e}")

# 验证异常详情中是否有我们期望的列
print("\n6. 验证异常详情列名")
try:
    abnormal_detail = pd.read_excel('result_with_string.xlsx', sheet_name='异常详情')
    print("异常详情包含的列：", abnormal_detail.columns.tolist())
except Exception as e:
    print(f"无法读取异常详情：{e}")

print("\n=== 结论 ===")
print("关键点：")
print("1. 如果 string_columns 真正有效，部门ID 应该从 int64 变为 object")
print("2. 异常详情中应该包含订单号（如果订单号在原始数据中）")
print("3. 订单号应该保持 'ORD-001' 这样的字符串格式")

# 清理
os.remove('direct_test.xlsx')
os.remove('result_no_string.xlsx')
os.remove('result_with_string.xlsx')

print("\n测试完成！")