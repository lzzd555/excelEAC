#!/usr/bin/env python3
"""
测试多列匹配的表合并
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


def test_multi_column():
    """测试多列匹配的表合并"""
    print("=== 测试3：多列匹配的表合并 ===\n")

    # 使用样例数据
    result = merge_excel_tables(
        table_a_file='tests/merge/sample_data/table_a_multi.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='tests/merge/sample_data/table_b_multi.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['订单号', '产品代码'],
        table_a_extra_columns=['客户', '数量'],
        table_b_extra_columns=['单价', '状态'],
        output_file='test_multi_column_result.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("\n合并结果:")
    print(result)

    if len(result) == 3:
        print("✅ 多列匹配测试通过")
        print(f"合并行数: {len(result)}")
        print(f"列名: {list(result.columns)}")
    else:
        print(f"❌ 多列匹配测试失败")
        print(f"期望: 3行匹配, 实际: {len(result)}行")


if __name__ == "__main__":
    test_multi_column()
