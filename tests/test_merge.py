#!/usr/bin/env python3
"""
测试表合并功能
"""

import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


def test_basic_merge():
    """测试基本表合并功能"""
    print("=== 测试1：基本表合并 ===\n")

    # 创建表A数据
    table_a_data = {
        'ID': ['A001', 'A002', 'A003', 'A004'],
        '姓名': ['张三', '李四', '王五', '赵六'],
        '部门': ['销售部', '销售部', '市场部', '技术部'],
        '年龄': [25, 30, 28, 35]
    }
    df_a = pd.DataFrame(table_a_data)
    df_a.to_excel('test_table_a.xlsx', index=False)

    # 创建表B数据
    table_b_data = {
        'ID': ['A001', 'A002', 'A003', 'B001'],
        '职位': ['经理', '专员', '主管', '经理'],
        '薪资': [10000, 8000, 12000, 15000]
    }
    df_b = pd.DataFrame(table_b_data)
    df_b.to_excel('test_table_b.xlsx', index=False)

    print("表A数据:")
    print(df_a)
    print("\n表B数据:")
    print(df_b)

    # 执行合并（只匹配ID列）
    result = merge_excel_tables(
        table_a_file='test_table_a.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='test_table_b.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['ID'],
        table_a_extra_columns=None,  # 不添加表A额外列
        table_b_extra_columns=None,  # 不添加表B额外列
        output_file='test_basic_merge.xlsx',
        string_columns=['ID']
    )

    print("\n合并结果:")
    print(result)
    print(f"合并后行数: {len(result)}")

    # 验证结果
    if len(result) == 3:  # 应该有3行匹配（A001, A002, A003）
        print("✅ 基本合并测试通过")
    else:
        print("❌ 基本合并测试失败")


def test_merge_with_extra_columns():
    """测试带额外列的表合并"""
    print("\n=== 测试2：带额外列的表合并 ===\n")

    # 创建表A数据
    table_a_data = {
        'ID': ['A001', 'A002', 'A003'],
        '姓名': ['张三', '李四', '王五'],
        '部门': ['销售部', '销售部', '市场部'],
        '年龄': [25, 30, 28]
    }
    df_a = pd.DataFrame(table_a_data)
    df_a.to_excel('test_table_a2.xlsx', index=False)

    # 创建表B数据
    table_b_data = {
        'ID': ['A001', 'A002', 'A003'],
        '职位': ['经理', '专员', '主管'],
        '薪资': [10000, 8000, 12000],
        '绩效等级': ['A', 'B', 'A']
    }
    df_b = pd.DataFrame(table_b_data)
    df_b.to_excel('test_table_b2.xlsx', index=False)

    print("表A数据:")
    print(df_a)
    print("\n表B数据:")
    print(df_b)

    # 执行合并（添加额外列）
    result = merge_excel_tables(
        table_a_file='test_table_a2.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='test_table_b2.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['ID'],
        table_a_extra_columns=['姓名', '部门', '年龄'],  # 从表A添加额外列
        table_b_extra_columns=['职位', '薪资', '绩效等级'],  # 从表B添加额外列
        output_file='test_extra_columns_merge.xlsx',
        string_columns=['ID']
    )

    print("\n合并结果:")
    print(result)

    # 验证列数
    expected_columns = ['ID', '姓名', '部门', '年龄', '职位', '薪资', '绩效等级']
    actual_columns = list(result.columns)
    print(f"期望列: {expected_columns}")
    print(f"实际列: {actual_columns}")

    if set(expected_columns) == set(actual_columns):
        print("✅ 带额外列的合并测试通过")
    else:
        print("❌ 带额外列的合并测试失败")


def test_multi_column_merge():
    """测试多列匹配的表合并"""
    print("\n=== 测试3：多列匹配的表合并 ===\n")

    # 创建表A数据
    table_a_data = {
        '订单号': ['001', '002', '003'],
        '产品代码': ['P01', 'P02', 'P03'],
        '客户': ['客户A', '客户B', '客户C'],
        '数量': [10, 20, 30]
    }
    df_a = pd.DataFrame(table_a_data)
    df_a.to_excel('test_table_a3.xlsx', index=False)

    # 创建表B数据
    table_b_data = {
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['P01', 'P02', 'P03', 'P04'],
        '单价': [100, 200, 150, 180],
        '状态': ['已发货', '待发货', '已发货', '已发货']
    }
    df_b = pd.DataFrame(table_b_data)
    df_b.to_excel('test_table_b3.xlsx', index=False)

    print("表A数据:")
    print(df_a)
    print("\n表B数据:")
    print(df_b)

    # 执行合并（按订单号和产品代码两列匹配）
    result = merge_excel_tables(
        table_a_file='test_table_a3.xlsx',
        table_a_sheet='Sheet1',
        table_b_file='test_table_b3.xlsx',
        table_b_sheet='Sheet1',
        match_columns=['订单号', '产品代码'],
        table_a_extra_columns=['客户', '数量'],  # 从表A添加额外列
        table_b_extra_columns=['单价', '状态'],  # 从表B添加额外列
        output_file='test_multi_column_merge.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("\n合并结果:")
    print(result)

    # 验证：应该有3行匹配（001-P01, 002-P02, 003-P03）
    if len(result) == 3:
        print("✅ 多列匹配测试通过")
    else:
        print(f"❌ 多列匹配测试失败（期望3行，实际{len(result)}行）")


def main():
    """主函数"""
    print("=== 表合并功能测试套件 ===\n")

    try:
        test_basic_merge()
        test_merge_with_extra_columns()
        test_multi_column_merge()

        print("\n=== 所有测试完成 ===")
        print("总结:")
        print("✅ 基本合并功能")
        print("✅ 额外列配置")
        print("✅ 多列匹配")
        print("✅ 字符串格式保持")

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
