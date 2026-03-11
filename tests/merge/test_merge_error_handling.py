#!/usr/bin/env python3
"""
测试错误处理功能
"""

import unittest
import sys
import os
import tempfile
import pandas as pd
from pathlib import Path

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


class TestMergeErrorHandling(unittest.TestCase):
    """测试错误处理"""

    def setUp(self):
        """创建测试数据"""
        self.temp_dir = tempfile.mkdtemp()

        # 创建有效的表A数据
        self.table_a_data = {
            'ID': ['1', '2', '3'],
            '姓名': ['张三', '李四', '王五'],
            '部门': ['销售部', '销售部', '市场部']
        }
        self.df_a = pd.DataFrame(self.table_a_data)
        self.table_a_path = os.path.join(self.temp_dir, 'table_a.xlsx')
        self.df_a.to_excel(self.table_a_path, index=False)

    def tearDown(self):
        """清理临时文件"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_missing_column_in_table_a(self):
        """测试表A中缺少匹配列"""
        # 创建表B数据，但没有表A需要的ID列
        # 注意：这里表A有ID列，所以这个测试不会触发错误
        # 正确的测试应该让表A缺少ID列
        table_b_data = {
            '员工编号': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        # 这个测试实际上会通过，因为表A有ID列，表B有员工编号列
        # 这是正确的列映射，不应该抛出异常
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': '员工编号'},  # 正确的列映射
            table_a_extra_columns=['姓名'],
            table_b_extra_columns=['薪资']
        )

        # 应该返回合并后的结果
        self.assertGreater(len(result), 0)

    def test_missing_column_in_table_b(self):
        """测试表B中缺少匹配列"""
        # 创建缺少ID列的表B
        table_b_data = {
            '姓名': ['张三', '李四', '王五'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        with self.assertRaises(ValueError) as context:
            merge_excel_tables(
                table_a_file=self.table_a_path,
                table_a_sheet='Sheet1',
                table_b_file=table_b_path,
                table_b_sheet='Sheet1',
                match_columns={'ID': '员工编号'},
                table_a_extra_columns=['姓名'],
                table_b_extra_columns=['薪资']
            )

        self.assertIn('表B中不存在匹配列', str(context.exception))

    def test_missing_extra_column_in_table_a(self):
        """测试表A中缺少额外列"""
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        with self.assertRaises(ValueError) as context:
            merge_excel_tables(
                table_a_file=self.table_a_path,
                table_a_sheet='Sheet1',
                table_b_file=table_b_path,
                table_b_sheet='Sheet1',
                match_columns=['ID'],
                table_a_extra_columns=['职位'],  # 表A中没有这个列
                table_b_extra_columns=['薪资']
            )

        self.assertIn('表A额外列不存在', str(context.exception))

    def test_missing_extra_column_in_table_b(self):
        """测试表B中缺少额外列"""
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        with self.assertRaises(ValueError) as context:
            merge_excel_tables(
                table_a_file=self.table_a_path,
                table_a_sheet='Sheet1',
                table_b_file=table_b_path,
                table_b_sheet='Sheet1',
                match_columns=['ID'],
                table_a_extra_columns=['姓名'],
                table_b_extra_columns=['职位']  # 表B中没有这个列
            )

        self.assertIn('表B额外列不存在', str(context.exception))

    def test_file_not_found_table_a(self):
        """测试表A文件不存在"""
        with self.assertRaises(Exception):
            merge_excel_tables(
                table_a_file='nonexistent_a.xlsx',
                table_a_sheet='Sheet1',
                table_b_file=self.table_a_path,
                table_b_sheet='Sheet1',
                match_columns=['ID']
            )

    def test_file_not_found_table_b(self):
        """测试表B文件不存在"""
        with self.assertRaises(Exception):
            merge_excel_tables(
                table_a_file=self.table_a_path,
                table_a_sheet='Sheet1',
                table_b_file='nonexistent_b.xlsx',
                table_b_sheet='Sheet1',
                match_columns=['ID']
            )

    def test_sheet_not_found(self):
        """测试工作表不存在"""
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        with self.assertRaises(Exception):
            merge_excel_tables(
                table_a_file=self.table_a_path,
                table_a_sheet='NonexistentSheet',  # 不存在的工作表
                table_b_file=table_b_path,
                table_b_sheet='Sheet1',
                match_columns=['ID']
            )

    def test_invalid_match_columns_type(self):
        """测试无效的 match_columns 类型"""
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        # 测试元组类型（应该失败）
        with self.assertRaises(AttributeError):
            merge_excel_tables(
                table_a_file=self.table_a_path,
                table_a_sheet='Sheet1',
                table_b_file=table_b_path,
                table_b_sheet='Sheet1',
                match_columns=('ID', '姓名'),  # 元组类型
            )

    def test_empty_match_columns_dict(self):
        """测试空的匹配列字典"""
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        # 空字典应该允许（返回空DataFrame）
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_path,
            table_b_sheet='Sheet1',
            match_columns={},  # 空字典
            table_a_extra_columns=['姓名'],
            table_b_extra_columns=['薪资']
        )

        # 结果应该是空DataFrame
        self.assertEqual(len(result), 0)

    def test_missing_column_in_table_a_real(self):
        """测试表A中确实缺少匹配列的真实情况"""
        # 创建表A数据，去掉ID列
        table_a_data_no_id = {
            '姓名': ['张三', '李四', '王五'],
            '部门': ['销售部', '销售部', '市场部']
        }
        df_a_no_id = pd.DataFrame(table_a_data_no_id)
        table_a_no_id_path = os.path.join(self.temp_dir, 'table_a_no_id.xlsx')
        df_a_no_id.to_excel(table_a_no_id_path, index=False)

        # 创建表B数据
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        with self.assertRaises(ValueError) as context:
            merge_excel_tables(
                table_a_file=table_a_no_id_path,
                table_a_sheet='Sheet1',
                table_b_file=table_b_path,
                table_b_sheet='Sheet1',
                match_columns={'ID': 'ID'},  # 表A没有ID列
                table_a_extra_columns=['姓名'],
                table_b_extra_columns=['薪资']
            )

        self.assertIn('表A中不存在匹配列', str(context.exception))

    def test_output_file_path(self):
        """测试输出文件路径"""
        table_b_data = {
            'ID': ['1', '2', '3'],
            '薪资': [5000, 6000, 7000]
        }
        df_b = pd.DataFrame(table_b_data)
        table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(table_b_path, index=False)

        output_path = os.path.join(self.temp_dir, 'output.xlsx')

        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_path,
            table_b_sheet='Sheet1',
            match_columns=['ID'],
            output_file=output_path
        )

        # 验证文件确实被创建
        self.assertTrue(os.path.exists(output_path))

        # 验证返回的DataFrame不为空
        self.assertGreater(len(result), 0)


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)