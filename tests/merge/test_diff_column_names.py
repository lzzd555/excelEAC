#!/usr/bin/env python3
"""
测试不同列名映射的表合并功能
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


def test_different_column_names():
    """测试不同列名映射的表合并功能"""
    print("=== 测试：不同列名映射的表合并 ===\n")

    # 使用不同列名的样例数据
    result = merge_excel_tables(
        table_a_file='tests/merge/sample_data/table_a_correct.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='tests/merge/sample_data/table_b_correct.xlsx',
        table_b_sheet='Sheet1',
        match_columns={'ID': '员工编号', '部门': '部门编码'},  # 列名映射
        table_a_extra_columns=['姓名', '职位', '入职日期'],  # 表A额外列
        table_b_extra_columns=['岗位', '薪资', '入职年份'],  # 表B额外列
        output_file='test_diff_column_names_result.xlsx',
        string_columns=['ID', '员工编号']
    )

    print("\n合并结果:")
    print(result)

    # 验证结果
    expected_columns = ['ID', '部门', '姓名', '职位', '入职日期', '岗位', '薪资', '入职年份']

    if len(result) == 3 and all(col in result.columns for col in expected_columns):
        print("✅ 不同列名映射测试通过")
        print(f"合并行数: {len(result)}")
        print(f"列数: {len(result.columns)}")
        print("列名:", list(result.columns))

        # 验证具体数据
        print("\n具体数据验证:")
        print("ID=001 的记录:", result[result['ID'] == '001'].iloc[0].to_dict())
        print("ID=002 的记录:", result[result['ID'] == '002'].iloc[0].to_dict())
        print("ID=003 的记录:", result[result['ID'] == '003'].iloc[0].to_dict())
    else:
        print("❌ 不同列名映射测试失败")
        print(f"期望行数: 3, 实际行数: {len(result)}")
        print(f"期望列数: {len(expected_columns)}, 实际列数: {len(result.columns)}")
        if result.columns is not None:
            print(f"缺少的列: {set(expected_columns) - set(result.columns)}")


def test_backward_compatibility():
    """测试向后兼容性：List[str] 格式仍然有效"""
    print("\n=== 测试：向后兼容性（List[str]格式）===\n")

    # 使用旧格式（List[str]）
    result = merge_excel_tables(
        table_a_file='tests/merge/sample_data/table_a_basic.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='tests/merge/sample_data/table_b_basic.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['ID'],  # 旧格式，应该自动转换为 {'ID': 'ID'}
        table_a_extra_columns=None,
        table_b_extra_columns=None,
        output_file='test_backward_compatibility_result.xlsx',
        string_columns=['ID']
    )

    print("\n合并结果:")
    print(result)

    # 验证结果
    if len(result) == 3 and 'ID' in result.columns:
        print("✅ 向后兼容性测试通过")
        print(f"合并行数: {len(result)}")
        print(f"列数: {len(result.columns)}")
    else:
        print("❌ 向后兼容性测试失败")


if __name__ == "__main__":
    test_different_column_names()
    test_backward_compatibility()
    print("\n所有测试完成！")