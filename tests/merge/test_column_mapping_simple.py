#!/usr/bin/env python3
"""
测试列名映射功能的单元测试（简化版）
"""

import unittest
import sys
import os
import tempfile
import pandas as pd

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


class TestColumnMappingSimple(unittest.TestCase):
    """测试列名映射功能 - 简化版"""

    def setUp(self):
        """创建测试数据"""
        self.temp_dir = tempfile.mkdtemp()

        # 创建表A数据
        df_a = pd.DataFrame({
            'ID': ['001', '002', '003', '004'],
            '姓名': ['张三', '李四', '王五', '赵六'],
            '部门': ['销售部', '销售部', '市场部', '技术部']
        })
        self.table_a_path = os.path.join(self.temp_dir, 'table_a.xlsx')
        df_a.to_excel(self.table_a_path, index=False)

        # 创建表B数据
        df_b = pd.DataFrame({
            '员工编号': ['001', '002', '003', '005'],
            '部门编码': ['销售部', '销售部', '市场部', '财务部'],
            '薪资': [5000, 6000, 7000, 9000]
        })
        self.table_b_path = os.path.join(self.temp_dir, 'table_b.xlsx')
        df_b.to_excel(self.table_b_path, index=False)

    def tearDown(self):
        """清理临时文件"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_basic_mapping(self):
        """测试基本映射功能"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': '员工编号'},
            table_a_extra_columns=['姓名', '部门'],
            table_b_extra_columns=['薪资'],
            output_file=os.path.join(self.temp_dir, 'result_basic.xlsx'),
            string_columns=['ID', '员工编号']
        )

        # 验证结果
        self.assertEqual(len(result), 3)  # 3条匹配记录
        expected_columns = ['ID', '部门', '姓名', '薪资']
        for col in expected_columns:
            self.assertIn(col, result.columns)

    def test_multi_column_mapping(self):
        """测试多列映射"""
        # 重新创建表B，添加部门列
        df_b = pd.DataFrame({
            '员工编号': ['001', '002', '003', '005'],
            '部门': ['销售部', '销售部', '市场部', '财务部'],  # 直接使用相同的部门名
            '薪资': [5000, 6000, 7000, 9000]
        })
        df_b.to_excel(self.table_b_path, index=False)

        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': '员工编号', '部门': '部门'},
            table_a_extra_columns=['姓名'],
            table_b_extra_columns=['薪资'],
            output_file=os.path.join(self.temp_dir, 'result_multi.xlsx')
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        expected_columns = ['ID', '部门', '姓名', '薪资']
        for col in expected_columns:
            self.assertIn(col, result.columns)

    def test_backward_compatibility(self):
        """测试向后兼容性"""
        # 为了测试向后兼容性，需要创建具有相同列名的表B
        df_b = pd.DataFrame({
            'ID': ['001', '002', '003', '005'],
            '部门': ['销售部', '销售部', '市场部', '财务部'],
            '薪资': [5000, 6000, 7000, 9000]
        })
        table_b_compat = os.path.join(self.temp_dir, 'table_b_compat.xlsx')
        df_b.to_excel(table_b_compat, index=False)

        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_compat,
            table_b_sheet='Sheet1',
            match_columns=['ID'],  # List格式
            table_a_extra_columns=['姓名'],
            table_b_extra_columns=['薪资'],
            output_file=os.path.join(self.temp_dir, 'result_backward.xlsx')
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertIn('ID', result.columns)
        self.assertIn('姓名', result.columns)
        self.assertIn('薪资', result.columns)

    def test_no_matches(self):
        """测试没有匹配记录"""
        # 创建没有匹配的表B
        df_b = pd.DataFrame({
            '员工编号': ['998', '999', '1000'],
            '部门': ['未知', '未知', '未知'],
            '薪资': [100, 200, 300]
        })
        table_b_no_match = os.path.join(self.temp_dir, 'table_b_no_match.xlsx')
        df_b.to_excel(table_b_no_match, index=False)

        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_no_match,
            table_b_sheet='Sheet1',
            match_columns={'ID': '员工编号'},
            table_a_extra_columns=['姓名'],
            table_b_extra_columns=['薪资'],
            output_file=os.path.join(self.temp_dir, 'result_no_match.xlsx')
        )

        # 验证结果
        self.assertEqual(len(result), 0)
        # 但列结构应该正确
        expected_columns = ['ID', '姓名', '薪资']
        for col in expected_columns:
            self.assertIn(col, result.columns)

    def test_empty_extra_columns(self):
        """测试空额外列"""
        result = merge_excel_tables(
            table_a_file=self.table_a_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': '员工编号'},
            table_a_extra_columns=[],  # 空列表
            table_b_extra_columns=[],  # 空列表
            output_file=os.path.join(self.temp_dir, 'result_empty.xlsx')
        )

        # 验证结果
        self.assertEqual(len(result), 3)
        self.assertEqual(len(result.columns), 1)  # 只有匹配列
        self.assertIn('ID', result.columns)


if __name__ == '__main__':
    unittest.main(verbosity=2)