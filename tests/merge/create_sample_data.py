#!/usr/bin/env python3
"""
创建表合并模块的测试样例数据
"""

import pandas as pd


def create_sample_data():
    """创建所有测试样例数据"""
    print("=== 创建表合并测试样例数据 ===\n")

    # 1. 创建基础测试数据（单列匹配）
    print("1. 创建基础测试数据...")

    table_a_basic_data = {
        'ID': ['A001', 'A002', 'A003', 'A004'],
        '姓名': ['张三', '李四', '王五', '赵六'],
        '部门': ['销售部', '销售部', '市场部', '技术部'],
        '年龄': [25, 30, 28, 35]
    }
    df_a_basic = pd.DataFrame(table_a_basic_data)

    table_b_basic_data = {
        'ID': ['A001', 'A002', 'A003', 'B001'],
        '职位': ['经理', '专员', '主管', '经理'],
        '薪资': [10000, 8000, 12000, 15000]
    }
    df_b_basic = pd.DataFrame(table_b_basic_data)

    # 保存到sample_data目录
    import os
    sample_dir = 'sample_data'
    os.makedirs(sample_dir, exist_ok=True)

    df_a_basic.to_excel(f'{sample_dir}/table_a_basic.xlsx', index=False)
    df_b_basic.to_excel(f'{sample_dir}/table_b_basic.xlsx', index=False)

    print(f"✅ 创建基础测试数据: {sample_dir}/table_a_basic.xlsx, {sample_dir}/table_b_basic.xlsx")

    # 2. 创建带额外列的测试数据
    print("\n2. 创建带额外列的测试数据...")

    table_a_extra_data = {
        'ID': ['A001', 'A002', 'A003'],
        '姓名': ['张三', '李四', '王五'],
        '部门': ['销售部', '销售部', '市场部'],
        '年龄': [25, 30, 28]
    }
    df_a_extra = pd.DataFrame(table_a_extra_data)

    table_b_extra_data = {
        'ID': ['A001', 'A002', 'A003'],
        '职位': ['经理', '专员', '主管'],
        '薪资': [10000, 8000, 12000],
        '绩效等级': ['A', 'B', 'A']
    }
    df_b_extra = pd.DataFrame(table_b_extra_data)

    df_a_extra.to_excel(f'{sample_dir}/table_a_extra.xlsx', index=False)
    df_b_extra.to_excel(f'{sample_dir}/table_b_extra.xlsx', index=False)

    print(f"✅ 创建带额外列的测试数据: {sample_dir}/table_a_extra.xlsx, {sample_dir}/table_b_extra.xlsx")

    # 3. 创建多列匹配测试数据
    print("\n3. 创建多列匹配测试数据...")

    table_a_multi_data = {
        '订单号': ['001', '002', '003'],
        '产品代码': ['P01', 'P02', 'P03'],
        '客户': ['客户A', '客户B', '客户C'],
        '数量': [10, 20, 30]
    }
    df_a_multi = pd.DataFrame(table_a_multi_data)

    table_b_multi_data = {
        '订单号': ['001', '002', '003', '004'],
        '产品代码': ['P01', 'P02', 'P03', 'P04'],
        '单价': [100, 200, 150, 180],
        '状态': ['已发货', '待发货', '已发货', '已发货']
    }
    df_b_multi = pd.DataFrame(table_b_multi_data)

    df_a_multi.to_excel(f'{sample_dir}/table_a_multi.xlsx', index=False)
    df_b_multi.to_excel(f'{sample_dir}/table_b_multi.xlsx', index=False)

    print(f"✅ 创建多列匹配测试数据: {sample_dir}/table_a_multi.xlsx, {sample_dir}/table_b_multi.xlsx")

    print("\n=== 测试样例数据创建完成 ===")
    print(f"\n所有样例数据保存在: {sample_dir}/")
    print(f"- table_a_basic.xlsx (基础测试表A)")
    print(f"- table_b_basic.xlsx (基础测试表B)")
    print(f"- table_a_extra.xlsx (带额外列表A)")
    print(f"- table_b_extra.xlsx (带额外列表B)")
    print(f"- table_a_multi.xlsx (多列匹配表A)")
    print(f"- table_b_multi.xlsx (多列匹配表B)")


if __name__ == "__main__":
    create_sample_data()
