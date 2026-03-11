#!/usr/bin/env python3
"""
运行合并模块的所有单元测试
"""

import unittest
import sys
import os
from pathlib import Path

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def run_all_tests():
    """运行所有合并相关的测试"""
    # 发现并运行测试
    loader = unittest.TestLoader()
    test_dir = Path(__file__).parent / 'tests' / 'merge'
    suite = loader.discover(str(test_dir), pattern='test_*.py')

    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)

    return result.wasSuccessful()

def run_specific_test(test_name):
    """运行特定的测试"""
    # 动态导入并运行测试
    if test_name == 'column_mapping':
        from tests.merge.test_column_mapping import TestColumnMapping
        suite = unittest.TestLoader().loadTestsFromTestCase(TestColumnMapping)
    elif test_name == 'parse_match_columns':
        from tests.merge.test_parse_match_columns import TestParseMatchColumns
        suite = unittest.TestLoader().loadTestsFromTestCase(TestParseMatchColumns)
    elif test_name == 'error_handling':
        from tests.merge.test_merge_error_handling import TestMergeErrorHandling
        suite = unittest.TestLoader().loadTestsFromTestCase(TestMergeErrorHandling)
    elif test_name == 'performance':
        from tests.merge.test_merge_performance import TestMergePerformance
        suite = unittest.TestLoader().loadTestsFromTestCase(TestMergePerformance)
    else:
        print(f"未知的测试: {test_name}")
        return False

    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)

    return result.wasSuccessful()

def main():
    """主函数"""
    print("=== Excel表合并功能 - 单元测试套件 ===\n")

    if len(sys.argv) > 1:
        # 运行特定测试
        test_name = sys.argv[1]
        print(f"运行特定测试: {test_name}")
        success = run_specific_test(test_name)
    else:
        # 运行所有测试
        print("运行所有测试...")
        success = run_all_tests()

    print("\n" + "=" * 50)
    if success:
        print("✅ 所有测试通过！")
        sys.exit(0)
    else:
        print("❌ 部分测试失败！")
        sys.exit(1)

if __name__ == '__main__':
    main()