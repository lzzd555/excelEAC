#!/usr/bin/env python3
"""
验证修复：确保字符串格式被正确保持
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation

def verify_fix():
    """验证修复后的功能"""
    print("=== 验证修复 ===\n")

    # 创建测试数据
    data = {
        '订单号': ['001', '002', '003', '004', '005'],
        '产品代码': ['A01', 'A02', 'A01', 'A03', 'A02'],
        '部门': ['销售部', '销售部', '市场部', '市场部', '销售部'],
        '月份': ['2024-01', '2024-01', '2024-01', '2024-02', '2024-02'],
        '计划值': [100, 200, 150, 120, 220],
        '实际值': [100, 210, 150, 120, 215]
    }

    df = pd.DataFrame(data)
    test_file = 'verify_test.xlsx'
    df.to_excel(test_file, index=False)

    print("测试数据：")
    print(df)

    print("\n运行验证...")
    result = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划值', '实际值'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='verify_result.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("\n分组结果：")
    print(result)

if __name__ == "__main__":
    verify_fix()
