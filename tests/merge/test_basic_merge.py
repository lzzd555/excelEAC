#!/usr/bin/env python3
"""
测试基本表合并功能（单列匹配）
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


def test_basic_merge():
    """测试基本表合并功能"""
    print("=== 测试1：基本表合并（单列匹配）===\n")

    # 使用样例数据
    result = merge_excel_tables(
        table_a_file='tests/merge/sample_data/table_a_basic.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='tests/merge/sample_data/table_b_basic.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['ID'],
        table_a_extra_columns=None,
        table_b_extra_columns=None,
        output_file='test_basic_merge_result.xlsx',
        string_columns=['ID']
    )

    print("\n合并结果:")
    print(result)

    # 验证结果
    if len(result) == 3 and 'ID' in result.columns:
        print("✅ 基本合并测试通过")
        print(f"合并行数: {len(result)}")
        print(f"列数: {len(result.columns)}")
    else:
        print("❌ 基本合并测试失败")


if __name__ == "__main__":
    test_basic_merge()
