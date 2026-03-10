#!/usr/bin/env python3
"""
最终验证：string_columns 在哪些地方真正有效
"""

import pandas as pd
from excel_validator import process_excel_with_validation
import os

def test_string_columns_effectiveness():
    """测试 string_columns 的实际效果"""

    print("=== string_columns 验证测试 ===\n")

    # 创建测试数据 - 重点是分组列可能是数字的情况
    test_data = {
        # 这些是原始数据列（不会出现在分组结果中）
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['A001', 'A002', 'A001', 'A003'],

        # 这些是分组列（可能会被转换为数字，然后我们需要保持为字符串）
        '部门代码': [1, 1, 2, 2],  # 数字格式
        '月份编号': [202401, 202401, 202402, 202402],  # 数字格式

        # 比较列
        '计划值': [100, 200, 150, 120],
        '实际值': [100, 200, 150, 120]  # 都正常
    }

    test_df = pd.DataFrame(test_data)
    test_df.to_excel('final_test.xlsx', index=False)

    print("1. 原始数据：")
    print(test_df)
    print("\n分组列原始类型：")
    print(f"部门代码: {test_df['部门代码'].dtype}")
    print(f"月份编号: {test_df['月份编号'].dtype}")

    print("\n2. 测试分组后的数据类型变化")

    # 测试：不使用 string_columns
    print("\n测试1：不使用 string_columns")
    result1 = process_excel_with_validation(
        input_file='final_test.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门代码', '月份编号'],
        compare_columns=['计划值', '实际值'],
        output_columns=['部门代码', '月份编号', '验证状态'],
        output_file='test1_no_string.xlsx'
    )

    print("分组结果（不使用 string_columns）：")
    print(result1)
    print("分组列类型：")
    print(f"部门代码: {result1['部门代码'].dtype}")
    print(f"月份编号: {result1['月份编号'].dtype}")

    # 测试：使用 string_columns
    print("\n测试2：使用 string_columns 保持分组列为字符串")
    result2 = process_excel_with_validation(
        input_file='final_test.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门代码', '月份编号'],
        compare_columns=['计划值', '实际值'],
        output_columns=['部门代码', '月份编号', '验证状态'],
        output_file='test2_with_string.xlsx',
        string_columns=['部门代码', '月份编号']  # 保持为字符串
    )

    print("分组结果（使用 string_columns）：")
    print(result2)
    print("分组列类型：")
    print(f"部门代码: {result2['部门代码'].dtype}")
    print(f"月份编号: {result2['月份编号'].dtype}")

    print("\n3. 测试异常详情中的 string_columns")
    # 创建有异常的数据
    abnormal_data = {
        '订单号': ['001', '002', '003'],
        '部门代码': [1, 1, 2],
        '月份编号': [202401, 202401, 202402],
        '计划值': [100, 200, 150],
        '实际值': [100, 210, 150]  # 第2行异常
    }
    abnormal_df = pd.DataFrame(abnormal_data)
    abnormal_df.to_excel('abnormal_test.xlsx', index=False)

    result_abnormal = process_excel_with_validation(
        input_file='abnormal_test.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门代码'],
        compare_columns=['计划值', '实际值'],
        output_columns=['部门代码', '验证状态'],
        output_file='abnormal_result.xlsx',
        string_columns=['部门代码', '订单号']
    )

    print("\n异常详情（应该包含订单号）：")
    try:
        abnormal_detail = pd.read_excel('abnormal_result.xlsx', sheet_name='异常详情')
        print(abnormal_detail)
        print("订单号类型:", abnormal_detail['订单号'].dtype if '订单号' in abnormal_detail.columns else "订单号列不存在")
    except Exception as e:
        print(f"读取异常详情失败：{e}")

    print("\n=== 结论 ===")
    print("✅ string_columns 在以下地方有效：")
    print("  1. 分组列：'1' → '1' (保持为字符串而不是数字)")
    print("  2. 异常详情中的原始数据列：'001' → '001' (保持前导零)")
    print("\n❌ string_columns 在以下地方无效：")
    print("  - 分组汇总结果中的原始数据列（如订单号、产品代码）")
    print("  - 因为分组汇总数据只包含分组列和统计信息")
    print("\n🎯 所以你的观察是正确的：订单号/产品代码确实不会出现在输出中")

    # 清理
    os.remove('final_test.xlsx')
    os.remove('test1_no_string.xlsx')
    os.remove('test2_with_string.xlsx')
    os.remove('abnormal_test.xlsx')
    os.remove('abnormal_result.xlsx')

if __name__ == "__main__":
    test_string_columns_effectiveness()