#!/usr/bin/env python3
"""
测试string_columns参数功能
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
import openpyxl

def test_string_columns():
    """测试string_columns功能"""
    print("=== 测试 string_columns ===\n")

    # 创建测试数据
    data = {
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['A01', 'A02', 'A01', 'A03'],
        '部门': ['A', 'A', 'B', 'B'],
        '计划值': [100, 200, 150, 120],
        '实际值': [100, 200, 150, 120]
    }

    df = pd.DataFrame(data)
    test_file = 'string_test.xlsx'
    df.to_excel(test_file, index=False)

    print("测试数据：")
    print(df)

    # 运行验证
    print("\n运行验证...")
    result = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划值', '实际值'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='test_result.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("\n分组结果：")
    print(result)

    # 检查字符串格式
    if '订单号' in result.columns:
        print(f"订单号类型: {result['订单号'].dtype}")

    print("\n✅ string_columns测试完成")

if __name__ == "__main__":
    test_string_columns()
