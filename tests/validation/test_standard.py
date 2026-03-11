#!/usr/bin/env python3
"""
标准测试用例：验证excel_validator.py的功能
基于用户修改后的test_abnormal_detail.py
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
import openpyxl

def standard_test():
    """标准测试用例"""
    print("=== 标准测试用例：验证excel_validator.py功能 ===\n")

    # 创建测试数据
    data = {
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['001', '002', '001', '003'],
        '部门': ['A', 'A', 'B', 'B'],
        '计划值': [100, 200, 150, 120],
        '实际值': [100, 210, 150, 120]
    }

    df = pd.DataFrame(data)
    test_file = 'standard_test.xlsx'
    df.to_excel(test_file, index=False)

    print("测试数据：")
    print(df)
    print("\n数据类型检查：")
    for col in df.columns:
        print(f"  {col}: {df[col].dtype}")

    # 运行验证
    print("\n运行验证...")
    result = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划值', '实际值'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='standard_test_output.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("\n分组结果：")
    print(result)

    print("\n测试完成！")

if __name__ == "__main__":
    standard_test()
