#!/usr/bin/env python3
"""
测试命令行参数解析函数
"""

import unittest
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import parse_match_columns


class TestParseMatchColumns(unittest.TestCase):
    """测试 parse_match_columns 函数"""

    def test_single_column_no_colon(self):
        """测试单个列，不带冒号（旧格式）"""
        result = parse_match_columns('ID')
        expected = {'ID': 'ID'}
        self.assertEqual(result, expected)

    def test_single_column_with_colon(self):
        """测试单个列，带冒号（新格式）"""
        result = parse_match_columns('ID:员工编号')
        expected = {'ID': '员工编号'}
        self.assertEqual(result, expected)

    def test_multiple_columns_no_colon(self):
        """测试多个列，不带冒号（旧格式）"""
        result = parse_match_columns('ID,部门')
        expected = {'ID': 'ID', '部门': '部门'}
        self.assertEqual(result, expected)

    def test_multiple_columns_with_colon(self):
        """测试多个列，带冒号（新格式）"""
        result = parse_match_columns('ID:员工编号,部门:部门编码')
        expected = {'ID': '员工编号', '部门': '部门编码'}
        self.assertEqual(result, expected)

    def test_mixed_format(self):
        """测试混合格式（部分带冒号，部分不带）"""
        result = parse_match_columns('ID:员工编号,部门')
        expected = {'ID': '员工编号', '部门': '部门'}
        self.assertEqual(result, expected)

    def test_with_spaces(self):
        """测试带空格的输入"""
        result = parse_match_columns('ID : 员工编号 , 部门 : 部门编码')
        expected = {'ID': '员工编号', '部门': '部门编码'}
        self.assertEqual(result, expected)

    def test_empty_input(self):
        """测试空输入"""
        result = parse_match_columns('')
        expected = {}
        self.assertEqual(result, expected)

    def test_only_commas(self):
        """测试只有逗号"""
        result = parse_match_columns(',,,')
        expected = {}  # 空字符串被跳过
        self.assertEqual(result, expected)

    def test_complex_mapping(self):
        """测试复杂的映射关系"""
        result = parse_match_columns('订单号:OrderID,客户名称:CustomerName,产品:Product,数量:Quantity,单价:UnitPrice')
        expected = {
            '订单号': 'OrderID',
            '客户名称': 'CustomerName',
            '产品': 'Product',
            '数量': 'Quantity',
            '单价': 'UnitPrice'
        }
        self.assertEqual(result, expected)

    def test_unicode_characters(self):
        """测试Unicode字符"""
        result = parse_match_columns('姓名:张三,年龄:18,邮箱:test@example.com')
        expected = {
            '姓名': '张三',
            '年龄': '18',
            '邮箱': 'test@example.com'
        }
        self.assertEqual(result, expected)

    def test_special_characters(self):
        """测试特殊字符"""
        result = parse_match_columns('col_1:value_1,col-2:value-2.col-3')
        expected = {
            'col_1': 'value_1',
            'col-2': 'value-2.col-3'
        }
        self.assertEqual(result, expected)


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)