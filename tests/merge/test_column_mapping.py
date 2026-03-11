#!/usr/bin/env python3
"""
测试列名映射功能的单元测试
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


class TestColumnMapping(unittest.TestCase):
    """测试列名映射功能"""

    def setUp(self):
        """创建测试数据"""
        # 创建临时目录
        self.temp_dir = tempfile.mkdtemp()

        # 创建表A数据
        self.table_a_data = {
            'id': ['1', '2', '3', '4'],  # 确保是字符串
            'name': ['张三', '李四', '王五', '赵六'],
            'dept': ['销售部', '销售部', '市场部', '技术部'],
            'salary': [5000, 6000, 7000, 8000]
        }

        # 确保数据类型正确
        self.df_a = pd.DataFrame(self.table_a_data)
        self.df_a['id'] = self.df_a['id'].astype(str)  # 确保是字符串
        self.df_a = pd.DataFrame(self.table_a_data)
        self.table_a_path = os.path.join(self.temp_dir, 'table_a.xlsx')
        self.df_a.to_excel(self.table_a_path, index=False)

        # 创建表B数据
        self.table_b_data = {
            'id': ['1', '2', '3', '5'],  # 添加id列用于向后兼容测试
            'employee_id': ['1', '2', '3', '5'],  # 确保是字符串
            'employee_name': ['张三', '李四', '王五', '钱七'],
            'dept': ['销售部', '销售部', '市场部', '财务部'],  # 添加dept列用于部分映射测试
            'department': ['销售部', '销售部', '市场部', '财务部'],
            'wage': [5000, 6000, 7000, 9000],
            'join_date': ['2020-01-01', '2020-02-01', '2020-03-01', '2021-01-01']
        }

        # 确保数据类型正确
        self.df_b = pd.DataFrame(self.table_b_data)
        self.df_b['id'] = self.df_b['id'].astype(str)
        self.df_b['employee_id'] = self.df_b['employee_id'].astype(str)
        self.df_b = pd.DataFrame(self.table_b_data)
        self.table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        self.df_b.to_excel(self.table_b_path, index=False)

    def tearDown(self):
        """清理临时文件"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_single_column_mapping(self):
        """测试单列映射"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'employee_id'},
            table_a_extra_columns=['name', 'dept'],
            table_b_extra_columns=['wage', 'join_date'],
            output_file=os.path.join(self.temp_dir, 'result_single.xlsx'),
            string_columns=['id', 'employee_id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)  # 3条匹配记录
        self.assertIn('id', result.columns)
        self.assertIn('name', result.columns)
        self.assertIn('wage', result.columns)

        # 验证数据正确性
        result_1 = result[result['id'] == '1']
        self.assertEqual(len(result_1), 1)
        self.assertEqual(result_1.iloc[0]['name'], '张三')
        self.assertEqual(result_1.iloc[0]['wage'], 5000)

    def test_multi_column_mapping(self):
        """测试多列映射"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'employee_id', 'dept': 'department'},
            table_a_extra_columns=['name'],
            table_b_extra_columns=['wage'],
            output_file=os.path.join(self.temp_dir, 'result_multi.xlsx'),
            string_columns=['id', 'employee_id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)  # 3条匹配记录
        self.assertIn('id', result.columns)
        self.assertIn('dept', result.columns)
        self.assertIn('name', result.columns)
        self.assertIn('wage', result.columns)

    def test_backward_compatibility_list(self):
        """测试向后兼容性：List[str]格式"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns=['id'],  # List格式，应该自动转换为Dict
            table_a_extra_columns=['name'],
            table_b_extra_columns=['wage'],
            output_file=os.path.join(self.temp_dir, 'result_list.xlsx'),
            string_columns=['id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertIn('id', result.columns)
        self.assertIn('name', result.columns)
        self.assertIn('wage', result.columns)

    def test_backward_compatibility_list_comma_separated(self):
        """测试向后兼容性：逗号分隔的List[str]"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns=['id', 'dept'],  # 多个列
            table_a_extra_columns=['name'],
            table_b_extra_columns=['wage'],
            output_file=os.path.join(self.temp_dir, 'result_multi_list.xlsx'),
            string_columns=['id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertIn('id', result.columns)
        self.assertIn('dept', result.columns)

    def test_no_extra_columns(self):
        """测试不包含额外列"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'employee_id'},
            table_a_extra_columns=None,
            table_b_extra_columns=None,
            output_file=os.path.join(self.temp_dir, 'result_no_extra.xlsx'),
            string_columns=['id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertEqual(len(result.columns), 1)  # 只有匹配列
        self.assertIn('id', result.columns)

    def test_empty_extra_columns(self):
        """测试空列表形式的额外列"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'employee_id'},
            table_a_extra_columns=[],
            table_b_extra_columns=[],
            output_file=os.path.join(self.temp_dir, 'result_empty_extra.xlsx'),
            string_columns=['id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertEqual(len(result.columns), 1)  # 只有匹配列
        self.assertIn('id', result.columns)

    def test_partial_mapping_with_same_names(self):
        """测试部分列名相同，部分不同的映射"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={
                'id': 'employee_id',  # 不同列名
                'dept': 'dept'         # 相同列名
            },
            table_a_extra_columns=['name'],
            table_b_extra_columns=['wage'],
            output_file=os.path.join(self.temp_dir, 'result_partial.xlsx'),
            string_columns=['id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertIn('id', result.columns)
        self.assertIn('dept', result.columns)
        self.assertIn('name', result.columns)
        self.assertIn('wage', result.columns)

    def test_string_columns_preservation(self):
        """测试字符串列保持功能"""
        # 创建包含前导零的测试数据
        table_a_data_with_leading_zeros = {
            'id': ['001', '002', '003'],
            'code': ['A01', 'B02', 'C03'],
            'name': ['Item1', 'Item2', 'Item3']
        }
        df_a_zeros = pd.DataFrame(table_a_data_with_leading_zeros)
        table_a_zeros_path = os.path.join(self.temp_dir, 'table_a_zeros.xlsx')
        df_a_zeros.to_excel(table_a_zeros_path, index=False)

        table_b_data_with_leading_zeros = {
            'item_id': ['001', '002', '004'],
            'item_code': ['A01', 'B02', 'C04'],
            'price': [100, 200, 300]
        }
        df_b_zeros = pd.DataFrame(table_b_data_with_leading_zeros)
        table_b_zeros_path = os.path.join(self.temp_dir, 'table_b_zeros.xlsx')
        df_b_zeros.to_excel(table_b_zeros_path, index=False)

        result = merge_excel_tables(
            table_a_file=table_a_zeros_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_zeros_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'item_id', 'code': 'item_code'},
            table_a_extra_columns=['name'],
            table_b_extra_columns=['price'],
            output_file=os.path.join(self.temp_dir, 'result_zeros.xlsx'),
            string_columns=['id', 'item_id', 'code', 'item_code']
        )

        # 验证前导零保持
        self.assertEqual(result.iloc[0]['id'], '001')
        self.assertEqual(result.iloc[1]['id'], '002')
        self.assertEqual(result.iloc[0]['code'], 'A01')

    def test_no_matches(self):
        """测试没有匹配记录的情况"""
        # 创建没有匹配的表B数据
        table_b_no_match = {
            'employee_id': ['99', '100', '101'],
            'department': ['财务部', '人事部', '研发部'],
            'wage': [10000, 11000, 12000]
        }
        df_b_no_match = pd.DataFrame(table_b_no_match)
        table_b_no_match_path = os.path.join(self.temp_dir, 'table_b_no_match.xlsx')
        df_b_no_match.to_excel(table_b_no_match_path, index=False)

        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_no_match_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'employee_id'},
            table_a_extra_columns=['name'],
            table_b_extra_columns=['wage'],
            output_file=os.path.join(self.temp_dir, 'result_no_match.xlsx'),
            string_columns=['id']
        )

        # 验证没有匹配记录
        self.assertEqual(len(result), 0)
        # 空结果但应该包含预期的列
        expected_columns = ['id', 'name', 'wage']
        for col in expected_columns:
            self.assertIn(col, result.columns)

    def test_duplicate_output_columns(self):
        """测试输出列重复的情况（额外列与匹配列同名）"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'id': 'employee_id'},
            table_a_extra_columns=['id', 'name'],  # id已经在匹配列中
            table_b_extra_columns=['wage'],
            output_file=os.path.join(self.temp_dir, 'result_duplicate.xlsx'),
            string_columns=['id']
        )

        # 验证结果
        self.assertEqual(len(result), 3)  # 应该有3条匹配记录
        # 验证没有重复列
        self.assertEqual(len(result.columns), 3)  # id只出现一次
        column_names = list(result.columns)
        self.assertEqual(column_names.count('id'), 1)
        self.assertIn('id', column_names)
        self.assertIn('name', column_names)
        self.assertIn('wage', column_names)


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)