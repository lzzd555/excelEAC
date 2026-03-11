#!/usr/bin/env python3
"""
测试带额外列的表合并
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


def test_extra_columns():
    """测试带额外列配置的表合并"""
    print("=== 测试2：带额外列的表合并 ===\n")

    # 使用样例数据
    result = merge_excel_tables(
        table_a_file='tests/merge/sample_data/table_a_extra.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='tests/merge/sample_data/table_b_extra.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['ID'],
        table_a_extra_columns=['姓名', '部门', '年龄'],
        table_b_extra_columns=['职位', '薪资', '绩效等级'],
        output_file='test_extra_columns_result.xlsx',
        string_columns=['ID']
    )

    print("\n合并结果:")
    print(result)

    expected_columns = ['ID', '姓名', '部门', '年龄', '职位', '薪资', '绩效等级']
    actual_columns = list(result.columns)

    print(f"\n期望列数: {len(expected_columns)}")
    print(f"实际列数: {len(actual_columns)}")

    if len(result) == 3 and set(expected_columns) == set(actual_columns):
        print("✅ 带额外列的合并测试通过")
        print(f"合并行数: {len(result)}")
        print(f"列名: {actual_columns}")
    else:
        print("❌ 带额外列的合并测试失败")


if __name__ == "__main__":
    test_extra_columns()
