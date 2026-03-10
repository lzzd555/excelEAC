#!/usr/bin/env python3
"""
测试 string_columns 参数的功能
验证数据格式保持是否正常工作
"""

import pandas as pd
from excel_validator import process_excel_with_validation
import os

def test_string_formatting():
    """测试字符串格式保持功能"""

    print("=== 测试 string_columns 参数 ===\n")

    # 创建测试数据 - 包含需要保持格式的列
    test_data = {
        # 这些列需要保持字符串格式
        '订单号': ['001', '002', '003', '001', '002', '003'],
        '产品代码': ['P001', 'P002', 'P003', 'P001', 'P002', 'P003'],
        '客户编号': ['C001', 'C002', 'C003', 'C001', 'C002', 'C003'],

        # 普通数值列
        '部门': [1, 1, 2, 2, 1, 2],  # 数字，但可能需要保持为字符串
        '计划数量': [100, 200, 150, 120, 180, 160],
        '实际数量': [100, 200, 150, 120, 180, 160],

        # 混合格式的列
        '备注': ['正常', '正常', '异常', '正常', '正常', '异常'],
        '日期': ['2024-01-01', '2024-01-01', '2024-01-01', '2024-02-01', '2024-02-01', '2024-02-01']
    }

    test_df = pd.DataFrame(test_data)

    # 保存原始数据
    test_df.to_excel('test_string_data.xlsx', index=False)

    print("1. 原始数据预览：")
    print("数据类型：")
    for col in test_df.columns:
        print(f"  {col}: {test_df[col].dtype}")
    print("\n数据内容：")
    print(test_df)
    print()

    # 测试1：不使用 string_columns
    print("2. 测试1：不使用 string_columns 参数")
    result1 = process_excel_with_validation(
        input_file='test_string_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='test_no_string.xlsx'
    )

    print("结果1（不使用string_columns）：")
    print("数据类型：")
    for col in result1.columns:
        print(f"  {col}: {result1[col].dtype}")
    print("数据内容：")
    print(result1)
    print()

    # 测试2：使用 string_columns 保持格式
    print("3. 测试2：使用 string_columns 参数")
    result2 = process_excel_with_validation(
        input_file='test_string_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='test_with_string.xlsx',
        string_columns=['部门']  # 保持部门为字符串格式
    )

    print("结果2（使用string_columns=['部门']）：")
    print("数据类型：")
    for col in result1.columns:
        print(f"  {col}: {result2[col].dtype}")
    print("数据内容：")
    print(result2)
    print()

    # 测试3：保持多个字符串列
    print("4. 测试3：保持多个字符串列")
    result3 = process_excel_with_validation(
        input_file='test_string_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='test_multiple_strings.xlsx',
        string_columns=['部门', '订单号', '产品代码']  # 多个列
    )

    print("结果3（使用string_columns=['部门', '订单号', '产品代码']）：")
    print("数据类型：")
    for col in result3.columns:
        print(f"  {col}: {result3[col].dtype}")
    print("数据内容：")
    print(result3)
    print()

    # 检查输出文件的内容
    print("5. 检查输出文件...")
    try:
        # 检查验证结果sheet
        df_check = pd.read_excel('test_with_string.xlsx', sheet_name='验证结果')
        print("输出文件内容：")
        print(df_check)
        print("数据类型：")
        for col in df_check.columns:
            print(f"  {col}: {df_check[col].dtype}")
    except Exception as e:
        print(f"读取输出文件失败：{e}")

    # 清理测试文件
    print("\n6. 清理测试文件...")
    test_files = ['test_string_data.xlsx', 'test_no_string.xlsx',
                 'test_with_string.xlsx', 'test_multiple_strings.xlsx']
    for file in test_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"已删除：{file}")

    print("\n=== 测试完成 ===")

def test_edge_cases():
    """测试边界情况"""
    print("\n=== 边界情况测试 ===\n")

    # 创建边界测试数据
    edge_data = {
        '零开头': ['001', '002', '003', '000', '010', '100'],
        '混合格式': ['1', '2', '3.14', '001', '002', '003'],
        '空值': ['A', 'B', None, 'D', 'E', 'F'],
        '正常数据': [10, 20, 30, 40, 50, 60]
    }

    edge_df = pd.DataFrame(edge_data)
    edge_df.to_excel('test_edge_data.xlsx', index=False)

    print("边界测试数据：")
    print(edge_df)
    print("数据类型：")
    for col in edge_df.columns:
        print(f"  {col}: {edge_df[col].dtype}")

    # 测试边界情况
    try:
        result = process_excel_with_validation(
            input_file='test_edge_data.xlsx',
            sheet_name='Sheet1',
            group_columns=['零开头'],
            compare_columns=['正常数据', '正常数据'],  # 相同的数据，确保正常
            output_columns=['零开头', '验证状态', '总行数'],
            output_file='test_edge_output.xlsx',
            string_columns=['零开头', '混合格式']
        )

        print("\n边界测试结果：")
        print(result)
        print("数据类型：")
        for col in result.columns:
            print(f"  {col}: {result[col].dtype}")

    except Exception as e:
        print(f"边界测试失败：{e}")

    # 清理
    if os.path.exists('test_edge_data.xlsx'):
        os.remove('test_edge_data.xlsx')
    if os.path.exists('test_edge_output.xlsx'):
        os.remove('test_edge_output.xlsx')

    print("\n=== 边界测试完成 ===")

def test_with_abnormal_data():
    """测试包含异常数据的情况"""
    print("\n=== 异常数据测试 ===\n")

    # 创建有异常的数据
    abnormal_data = {
        '订单号': ['001', '002', '003', '004', '005'],
        '产品代码': ['P001', 'P002', 'P003', 'P004', 'P005'],
        '部门': ['A', 'A', 'B', 'B', 'C'],
        '计划数量': [100, 200, 150, 120, 180],
        '实际数量': [100, 200, 150, 125, 180],  # 第4行有异常
        '状态': ['正常', '正常', '正常', '异常', '正常']
    }

    abnormal_df = pd.DataFrame(abnormal_data)
    abnormal_df.to_excel('test_abnormal_data.xlsx', index=False)

    print("异常测试数据：")
    print(abnormal_df)

    # 测试异常数据下的string_columns
    result = process_excel_with_validation(
        input_file='test_abnormal_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='test_abnormal_output.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("\n异常数据测试结果：")
    print(result)

    # 检查异常详情sheet
    try:
        abnormal_detail = pd.read_excel('test_abnormal_output.xlsx', sheet_name='异常详情')
        print("\n异常详情：")
        print(abnormal_detail)
    except:
        print("\n没有异常详情sheet")

    # 清理
    if os.path.exists('test_abnormal_data.xlsx'):
        os.remove('test_abnormal_data.xlsx')
    if os.path.exists('test_abnormal_output.xlsx'):
        os.remove('test_abnormal_output.xlsx')

    print("\n=== 异常数据测试完成 ===")

if __name__ == "__main__":
    # 运行所有测试
    test_string_formatting()
    test_edge_cases()
    test_with_abnormal_data()
    print("\n🎉 所有测试完成！")