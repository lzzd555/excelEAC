#!/usr/bin/env python3
"""
直接测试：创建一个能清晰展示 string_columns 效果的用例
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation

def simple_test():
    print("=== 直接测试 string_columns 效果 ===\n")

    # 创建测试数据 - 确保某些列是数字格式
    test_data = {
        '部门ID': [1, 1, 2, 2, 3, 3],  # int64
        '客户ID': [101, 102, 103, 104, 105, 106],  # int64
        '订单金额': [1000, 2000, 1500, 3000, 2500, 1800],  # int64
        '实际支付': [1000, 2000, 1500, 3100, 2500, 1800]  # int64
    }

    test_df = pd.DataFrame(test_data)
    test_df.to_excel('direct_test.xlsx', index=False)

    print("1. 原始数据预览：")
    print(test_df)
    print("数据类型：")
    for col in test_df.columns:
        print(f"  {col}: {test_df[col].dtype}")

    print("\n2. 测试：不使用 string_columns")
    result1 = process_excel_with_validation(
        input_file='direct_test.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门ID'],
        compare_columns=['订单金额', '实际支付'],
        output_columns=['部门ID', '验证状态', '总行数'],
        output_file='result_no_string.xlsx'
    )

    print("结果（不使用 string_columns）：")
    print(result1)

    print("\n3. 测试：使用 string_columns 将部门ID和客户ID转为字符串")
    result2 = process_excel_with_validation(
        input_file='direct_test.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门ID'],
        compare_columns=['订单金额', '实际支付'],
        output_columns=['部门ID', '验证状态', '总行数'],
        output_file='result_with_string.xlsx',
        string_columns=['部门ID', '客户ID']
    )

    print("结果（使用 string_columns=['部门ID', '客户ID']）：")
    print(result2)

    print("\n4. 验证异常详情...")
    try:
        abnormal_detail = pd.read_excel('result_with_string.xlsx', sheet_name='异常详情')
        print("异常详情包含的列：", list(abnormal_detail.columns))
    except Exception as e:
        print(f"读取异常详情失败：{e}")

if __name__ == "__main__":
    simple_test()
