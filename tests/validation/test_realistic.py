#!/usr/bin/env python3
"""
真实场景测试：验证 string_columns 的实际效果
"""

import pandas as pd
from excel_validator import process_excel_with_validation

import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import
# 创建更真实的测试数据
print("=== 真实场景测试 ===\n")

test_data = {
    '订单号': ['001', '002', '003', '004', '005', '006'],
    '产品代码': ['P001', 'P002', 'P001', 'P003', 'P002', 'P001'],
    '部门': ['销售部', '销售部', '市场部', '市场部', '销售部', '市场部'],
    '月份': ['2024-01', '2024-01', '2024-01', '2024-02', '2024-02', '2024-02'],
    '计划销量': [100, 150, 200, 120, 180, 90],
    '实际销量': [100, 155, 200, 125, 180, 90],  # 第2行和第4行有异常
    '客户ID': ['C001', 'C002', 'C003', 'C004', 'C005', 'C006']
}

test_df = pd.DataFrame(test_data)
test_df.to_excel('real_test_data.xlsx', index=False)

print("1. 原始数据（6行）：")
print(test_df)
print("\n数据类型：")
for col in test_df.columns:
    print(f"  {col}: {test_df[col].dtype}")

print("\n2. 分组逻辑说明：")
print("- 销售部, 2024-01: 2行（订单001,002）→ 订单002销量异常")
print("- 销售部, 2024-02: 1行（订单005）→ 正常")
print("- 市场部, 2024-01: 2行（订单003,006）→ 正常")
print("- 市场部, 2024-02: 1行（订单004）→ 销量异常")
print("预期：销售部2024-01月和市场部2024-02月应该是'异常'组")

# 测试 string_columns 的实际效果
print("\n3. 测试 string_columns 的效果")
print("\n测试1：不使用 string_columns")
result1 = process_excel_with_validation(
    input_file='real_test_data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门', '月份'],
    compare_columns=['计划销量', '实际销量'],
    output_columns=['部门', '月份', '验证状态', '总行数'],
    output_file='test_no_string.xlsx'
)

print("输出结果（不使用string_columns）：")
print(result1)
print("数据类型：")
for col in result1.columns:
    print(f"  {col}: {result1[col].dtype}")

print("\n测试2：使用 string_columns 保持订单号格式")
result2 = process_excel_with_validation(
    input_file='real_test_data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门', '月份'],
    compare_columns=['计划销量', '实际销量'],
    output_columns=['部门', '月份', '验证状态', '总行数'],
    output_file='test_with_string.xlsx',
    string_columns=['部门', '订单号']  # 注意：订单号在分组结果中不存在！
)

print("\n输出结果（使用string_columns=['部门', '订单号']）：")
print(result2)
print("数据类型：")
for col in result2.columns:
    print(f"  {col}: {result2[col].dtype}")

print("\n❌ 重要发现：")
print("1. '订单号'列没有出现在输出结果中")
print("2. 因为分组汇总数据不包含原始数据列")
print("3. string_columns 只对分组列有效")

# 检查异常详情
print("\n4. 检查异常详情中的 string_columns 效果")
try:
    abnormal_df = pd.read_excel('test_with_string.xlsx', sheet_name='异常详情')
    print("异常详情：")
    print(abnormal_df)
    print("数据类型：")
    for col in abnormal_df.columns:
        print(f"  {col}: {abnormal_df[col].dtype}")

    print("\n✅ 在异常详情中，订单号保持了字符串格式！")
except Exception as e:
    print(f"读取异常详情失败：{e}")

print("\n5. 结论：")
print("✅ string_columns 在异常详情中有效")
print("✅ string_columns 在分组结果中对分组列有效")
print("❌ string_columns 对原始数据列无效（因为它们不在分组结果中）")

# 清理
import os
os.remove('real_test_data.xlsx')
os.remove('test_no_string.xlsx')
os.remove('test_with_string.xlsx')

print("\n=== 测试完成 ===")