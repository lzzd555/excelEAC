#!/usr/bin/env python3
"""
创建测试数据，用于测试不同列名映射的表合并功能
"""

import pandas as pd
import os

# 确保测试数据目录存在
os.makedirs('tests/merge/sample_data', exist_ok=True)

# 创建表A数据：使用列名 ID, 姓名, 部门, 职位
table_a_data = {
    'ID': ['001', '002', '003', '004'],
    '姓名': ['张三', '李四', '王五', '赵六'],
    '部门': ['销售部', '销售部', '市场部', '技术部'],
    '职位': ['经理', '专员', '主管', '工程师'],
    '入职日期': ['2020-01-01', '2020-02-01', '2020-03-01', '2020-04-01']
}
df_a = pd.DataFrame(table_a_data)
df_a.to_excel('tests/merge/sample_data/table_a_diff_name.xlsx', index=False)
print("✅ 创建表A数据（不同列名映射测试）: table_a_diff_name.xlsx")

# 创建表B数据：使用列名 员工编号, 部门编码, 岗位, 薪资, 评级
table_b_data = {
    '员工编号': ['001', '002', '003', '005'],  # 004在表B中不存在，用于测试不匹配情况
    '部门编码': ['XS', 'XS', 'SC', 'JS'],   # XS=销售部, SC=市场部, JS=技术部
    '岗位': ['经理', '专员', '主管', '架构师'],
    '薪资': [15000, 8000, 12000, 20000],
    '入职年份': [2020, 2020, 2020, 2021]
}
df_b = pd.DataFrame(table_b_data)
df_b.to_excel('tests/merge/sample_data/table_b_diff_name.xlsx', index=False)
print("✅ 创建表B数据（不同列名映射测试）: table_b_diff_name.xlsx")

print("\n测试数据说明:")
print("- 表A列: ID, 姓名, 部门, 职位, 入职日期")
print("- 表B列: 员工编号, 部门编码, 岗位, 薪资, 入职年份")
print("- 映射关系: ID → 员工编号, 部门 → 部门编码")
print("- 预期结果: 3条匹配记录（ID=001,002,003）")