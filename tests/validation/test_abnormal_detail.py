#!/usr/bin/env python3
"""
用户的标准测试用例，测试异常详情中的字符串格式保持
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
import openpyxl

def test_abnormal_detail():
    """测试异常详情中的字符串格式"""
    print("=== 测试异常详情中的 string_columns ===\n")

    # 创建有异常的数据
    data = {
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['001', '002', '001', '003'],
        '部门': ['A', 'A', 'B', 'B'],
        '计划值': [100, 200, 150, 120],
        '实际值': [100, 210, 150, 120]  # 第2行有异常
    }

    df = pd.DataFrame(data)
    test_file = 'abnormal_test.xlsx'
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
        output_file='test_output.xlsx',
        string_columns=['订单号', '产品代码']  # 应该包含在异常详情中
    )

    print("\n分组结果：")
    print(result)

if __name__ == "__main__":
    test_abnormal_detail()
