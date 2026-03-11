#!/usr/bin/env python3
"""
测试辅助函数，用于创建测试数据
"""

import pandas as pd
import tempfile
import os


def create_merge_test_data(temp_dir=None, use_string_id=True):
    """
    创建用于合并测试的数据

    Args:
        temp_dir: 临时目录路径，如果为None则使用当前目录
        use_string_id: 是否使用字符串格式的ID

    Returns:
        tuple: (table_a_path, table_b_path)
    """
    if temp_dir is None:
        temp_dir = tempfile.mkdtemp()

    # 创建表A数据
    if use_string_id:
        table_a_data = {
            'ID': ['001', '002', '003', '004'],
            '姓名': ['张三', '李四', '王五', '赵六'],
            '部门': ['销售部', '销售部', '市场部', '技术部'],
            '薪资': [5000, 6000, 7000, 8000]
        }
    else:
        table_a_data = {
            'ID': [1, 2, 3, 4],
            '姓名': ['张三', '李四', '王五', '赵六'],
            '部门': ['销售部', '销售部', '市场部', '技术部'],
            '薪资': [5000, 6000, 7000, 8000]
        }

    df_a = pd.DataFrame(table_a_data)

    # 创建表B数据
    if use_string_id:
        table_b_data = {
            '员工编号': ['001', '002', '003', '005'],
            '员工姓名': ['张三', '李四', '王五', '钱七'],
            '部门编码': ['销售部', '销售部', '市场部', '财务部'],
            '工资': [5000, 6000, 7000, 9000],
            '入职日期': ['2020-01-01', '2020-02-01', '2020-03-01', '2021-01-01']
        }
    else:
        table_b_data = {
            '员工编号': [1, 2, 3, 5],
            '员工姓名': ['张三', '李四', '王五', '钱七'],
            '部门编码': ['销售部', '销售部', '市场部', '财务部'],
            '工资': [5000, 6000, 7000, 9000],
            '入职日期': ['2020-01-01', '2020-02-01', '2020-03-01', '2021-01-01']
        }

    df_b = pd.DataFrame(table_b_data)

    # 保存到文件
    table_a_path = os.path.join(temp_dir, 'table_a.xlsx')
    table_b_path = os.path.join(temp_dir, 'table_b.xlsx')

    # 使用dtype参数确保数据类型正确
    string_columns = ['ID', '员工编号'] if use_string_id else None

    df_a.to_excel(table_a_path, index=False)
    df_b.to_excel(table_b_path, index=False)

    return table_a_path, table_b_path


def create_no_match_test_data(temp_dir):
    """创建用于测试无匹配记录的数据"""
    # 表A数据
    table_a_data = {
        'ID': ['1', '2', '3'],
        'Name': ['A', 'B', 'C']
    }
    df_a = pd.DataFrame(table_a_data)
    table_a_path = os.path.join(temp_dir, 'table_a_no_match.xlsx')
    df_a.to_excel(table_a_path, index=False)

    # 表B数据（没有匹配的ID）
    table_b_data = {
        'EmployeeID': ['4', '5', '6'],
        'Dept': ['X', 'Y', 'Z']
    }
    df_b = pd.DataFrame(table_b_data)
    table_b_path = os.path.join(temp_dir, 'table_b_no_match.xlsx')
    df_b.to_excel(table_b_path, index=False)

    return table_a_path, table_b_path