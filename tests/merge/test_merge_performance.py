#!/usr/bin/env python3
"""
测试性能相关功能
"""

import unittest
import sys
import os
import tempfile
import time
import pandas as pd

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables


class TestMergePerformance(unittest.TestCase):
    """测试性能"""

    def setUp(self):
        """创建测试数据"""
        self.temp_dir = tempfile.mkdtemp()

        # 创建大量数据的测试用例
        large_data_size = 1000
        self.large_table_a_data = {
            'ID': [f'{i:04d}' for i in range(1, large_data_size + 1)],
            'Name': [f'User_{i}' for i in range(1, large_data_size + 1)],
            'Department': ['销售部' if i % 2 == 0 else '市场部' for i in range(large_data_size)],
            'Value': [i * 100 for i in range(large_data_size)]
        }
        self.df_a_large = pd.DataFrame(self.large_table_a_data)
        self.table_a_large_path = os.path.join(self.temp_dir, 'large_table_a.xlsx')
        self.df_a_large.to_excel(self.table_a_large_path, index=False)

        # 创建匹配的表B数据（只有一半匹配）
        match_count = large_data_size // 2
        self.large_table_b_data = {
            'EmployeeID': [f'{i:04d}' for i in range(1, match_count + 1)],
            'Department': ['销售部' if i % 2 == 0 else '市场部' for i in range(match_count)],
            'Salary': [i * 1000 for i in range(match_count)],
            'Bonus': [i * 100 for i in range(match_count)]
        }
        self.df_b_large = pd.DataFrame(self.large_table_b_data)
        self.table_b_large_path = os.path.join(self.temp_dir, 'large_table_b.xlsx')
        self.df_b_large.to_excel(self.table_b_large_path, index=False)

    def tearDown(self):
        """清理临时文件"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_merge_performance_with_dict_mapping(self):
        """测试使用字典映射的性能"""
        start_time = time.time()
        result = merge_excel_tables(
            table_a_file=self.table_a_large_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_large_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': 'EmployeeID', 'Department': 'Department'},
            table_a_extra_columns=['Name', 'Value'],
            table_b_extra_columns=['Salary', 'Bonus'],
            output_file=os.path.join(self.temp_dir, 'perf_dict_result.xlsx')
        )
        end_time = time.time()

        # 验证结果
        self.assertEqual(len(result), len(self.df_b_large))
        self.assertLess(end_time - start_time, 5.0)  # 应该在5秒内完成

        print(f"字典映射耗时: {end_time - start_time:.2f}秒")

    def test_merge_performance_with_list_mapping(self):
        """测试使用列表映射的性能（向后兼容）"""
        start_time = time.time()
        result = merge_excel_tables(
            table_a_file=self.table_a_large_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_large_path,
            table_b_sheet='Sheet1',
            match_columns=['ID', 'Department'],  # 列表格式
            table_a_extra_columns=['Name', 'Value'],
            table_b_extra_columns=['Salary', 'Bonus'],
            output_file=os.path.join(self.temp_dir, 'perf_list_result.xlsx')
        )
        end_time = time.time()

        # 验证结果
        self.assertEqual(len(result), len(self.df_b_large))
        self.assertLess(end_time - start_time, 5.0)  # 应该在5秒内完成

        print(f"列表映射耗时: {end_time - start_time:.2f}秒")

    def test_single_column_vs_multi_column(self):
        """测试单列匹配 vs 多列匹配的性能"""
        # 创建小一点的测试数据（500条）
        small_size = 500
        small_table_a = {
            'ID': [f'{i:04d}' for i in range(small_size)],
            'Dept': ['A' if i % 3 == 0 else 'B' for i in range(small_size)],
            'Name': [f'User_{i}' for i in range(small_size)]
        }
        df_a_small = pd.DataFrame(small_table_a)
        table_a_small_path = os.path.join(self.temp_dir, 'small_table_a.xlsx')
        df_a_small.to_excel(table_a_small_path, index=False)

        small_table_b = {
            'EmployeeID': [f'{i:04d}' for i in range(small_size // 2)],
            'Department': ['A' if i % 3 == 0 else 'B' for i in range(small_size // 2)],
            'Salary': [i * 100 for i in range(small_size // 2)]
        }
        df_b_small = pd.DataFrame(small_table_b)
        table_b_small_path = os.path.join(self.temp_dir, 'small_table_b.xlsx')
        df_b_small.to_excel(table_b_small_path, index=False)

        # 测试单列匹配
        start_time = time.time()
        result_single = merge_excel_tables(
            table_a_file=table_a_small_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_small_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': 'EmployeeID'},
            output_file=os.path.join(self.temp_dir, 'single_col_result.xlsx')
        )
        single_time = time.time() - start_time

        # 测试多列匹配
        start_time = time.time()
        result_multi = merge_excel_tables(
            table_a_file=table_a_small_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_small_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': 'EmployeeID', 'Dept': 'Department'},
            output_file=os.path.join(self.temp_dir, 'multi_col_result.xlsx')
        )
        multi_time = time.time() - start_time

        # 验证结果
        self.assertEqual(len(result_single), len(df_b_small))
        self.assertEqual(len(result_multi), len(df_b_small) // 3)  # 多列匹配应该减少匹配数

        print(f"单列匹配耗时: {single_time:.2f}秒")
        print(f"多列匹配耗时: {multi_time:.2f}秒")

    def test_large_dataset_memory_usage(self):
        """测试大数据集的内存使用"""
        # 监控内存使用情况（这只是一个简单的测试）
        import psutil
        process = psutil.Process()
        initial_memory = process.memory_info().rss / 1024 / 1024  # MB

        # 执行合并操作
        result = merge_excel_tables(
            table_a_file=self.table_a_large_path,
            table_a_sheet='Sheet1',
            table_b_file=self.table_b_large_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': 'EmployeeID'},
            output_file=os.path.join(self.temp_dir, 'memory_test.xlsx')
        )

        final_memory = process.memory_info().rss / 1024 / 1024  # MB
        memory_increase = final_memory - initial_memory

        # 内存增加不应该超过100MB（这个值可以根据实际情况调整）
        self.assertLess(memory_increase, 100)

        print(f"内存使用增加: {memory_increase:.2f}MB")
        print(f"合并结果行数: {len(result)}")

    def test_string_columns_performance(self):
        """测试字符串列处理性能"""
        # 创建带有字符串列的数据
        string_data_size = 1000
        string_table_a = {
            'ID': [f'{i:04d}' for i in range(string_data_size)],
            'Code': [f'A{i:03d}' for i in range(string_data_size)],
            'Name': [f'Product_{i}' for i in range(string_data_size)],
            'Desc': [f'Description for product {i}' for i in range(string_data_size)]
        }
        df_a_string = pd.DataFrame(string_table_a)
        table_a_string_path = os.path.join(self.temp_dir, 'string_table_a.xlsx')
        df_a_string.to_excel(table_a_string_path, index=False)

        # 创建匹配的表B
        match_count = string_data_size // 2
        string_table_b = {
            'ItemID': [f'{i:04d}' for i in range(1, match_count + 1)],
            'ItemCode': [f'A{i:03d}' for i in range(match_count)],
            'Price': [i * 10 for i in range(match_count)]
        }
        df_b_string = pd.DataFrame(string_table_b)
        table_b_string_path = os.path.join(self.temp_dir, 'string_table_b.xlsx')
        df_b_string.to_excel(table_b_string_path, index=False)

        # 测试带字符串列的合并
        start_time = time.time()
        result = merge_excel_tables(
            table_a_file=table_a_string_path,
            table_a_sheet='Sheet1',
            table_b_file=table_b_string_path,
            table_b_sheet='Sheet1',
            match_columns={'ID': 'ItemID', 'Code': 'ItemCode'},
            table_a_extra_columns=['Name', 'Desc'],
            table_b_extra_columns=['Price'],
            output_file=os.path.join(self.temp_dir, 'string_perf_result.xlsx'),
            string_columns=['ID', 'ItemID', 'Code', 'ItemCode']
        )
        end_time = time.time()

        # 验证前导零保持
        self.assertEqual(result.iloc[0]['ID'], '0001')
        self.assertEqual(result.iloc[0]['Code'], 'A000')

        print(f"字符串列合并耗时: {end_time - start_time:.2f}秒")


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)