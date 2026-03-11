#!/usr/bin/env python3
"""
最终验证：string_columns 在哪些地方真正有效
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation

def test_final():
    """测试 string_columns 的实际效果"""
    print("=== 最终验证：string_columns 在哪些地方真正有效 ===\n")

    # 创建测试数据 - 重点是分组列可能是数字的情况
    test_data = {
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['A01', 'A02', 'A01', 'A03'],
        '部门代码': [1, 1, 2, 2],  # 数字格式
        '月份编号': [202401, 202401, 202402, 202402],  # 数字格式
        '计划金额': [1000, 2000, 1500, 1200],  # int64
        '实际金额': [1000, 2000, 1500, 1200]  # int64
    }

    test_df = pd.DataFrame(test_data)
    test_file = 'final_test.xlsx'
    test_df.to_excel(test_file, index=False)

    print("1. 原始数据：")
    print(test_df)
    print("分组列原始类型：")
    for col in ['部门代码', '月份编号']:
        print(f"  {col}: {test_df[col].dtype}")

    print("\n2. 测试分组后的数据类型变化")

    # 测试：不使用 string_columns
    print("测试1：不使用 string_columns")
    result1 = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门代码', '月份编号'],
        compare_columns=['计划金额', '实际金额'],
        output_columns=['部门代码', '月份编号', '验证状态'],
        output_file='test1_no_string.xlsx'
    )

    print("分组结果（不使用 string_columns）：")
    print(result1)
    print("分组列类型：")
    for col in ['部门代码', '月份编号']:
        print(f"  {col}: {result1[col].dtype}")

    # 测试：使用 string_columns
    print("\n测试2：使用 string_columns 保持分组列为字符串")
    result2 = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门代码', '月份编号'],
        compare_columns=['计划金额', '实际金额'],
        output_columns=['部门代码', '月份编号', '验证状态'],
        output_file='test2_with_string.xlsx',
        string_columns=['部门代码', '月份编号']
    )

    print("分组结果（使用 string_columns）：")
    print(result2)
    print("分组列类型：")
    for col in ['部门代码', '月份编号']:
        print(f"  {col}: {result2[col].dtype}")

if __name__ == "__main__":
    test_final()
