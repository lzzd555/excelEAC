#!/usr/bin/env python3
"""
运行所有测试的脚本
"""

import subprocess
import sys
import os


def run_test(test_file, description):
    """运行单个测试"""
    print(f"\n{'='*60}")
    print(f"运行测试: {description}")
    print(f"{'='*60}")
    try:
        # 从项目根目录运行
        result = subprocess.run(
            [sys.executable, test_file],
            capture_output=True,
            text=True,
            timeout=60
        )
        print(result.stdout)
        if result.returncode != 0:
            print(f"❌ 测试失败: {result.stderr}")
            return False
        print(f"✅ 测试通过")
        return True
    except subprocess.TimeoutExpired:
        print(f"❌ 测试超时")
        return False
    except Exception as e:
        print(f"❌ 测试执行失败: {e}")
        return False


def main():
    """主函数"""
    tests_dir = 'tests'
    project_root = os.path.dirname(os.path.abspath(__file__))

    # 验证模块测试
    validation_tests = [
        (f'{tests_dir}/validation/test_standard.py', '标准验证测试'),
        (f'{tests_dir}/validation/test_abnormal_detail.py', '异常详情测试'),
        (f'{tests_dir}/validation/test_string_columns.py', '字符串列测试'),
    ]

    # 合并模块测试
    merge_tests = [
        (f'{tests_dir}/merge/test_basic_merge.py', '基本合并测试'),
        (f'{tests_dir}/merge/test_extra_columns.py', '带额外列的合并测试'),
        (f'{tests_dir}/merge/test_multi_column.py', '多列匹配测试'),
    ]

    all_tests = validation_tests + merge_tests

    print("=== Excel工具包 - 测试套件 ===")
    print(f"将运行 {len(all_tests)} 个测试\n")

    results = []
    for test_file, description in all_tests:
        success = run_test(test_file, description)
        results.append((test_file, description, success))

    # 汇总结果
    print(f"\n{'='*60}")
    print("测试结果汇总")
    print(f"{'='*60}")

    passed = sum(1 for _, _, success in results if success)
    failed = sum(1 for _, _, success in results if not success)

    print(f"通过: {passed}/{len(all_tests)}")
    print(f"失败: {failed}/{len(all_tests)}")

    print("\n详细结果:")
    for test_file, description, success in results:
        status = "✅ 通过" if success else "❌ 失败"
        print(f"  {status} - {description} ({test_file})")

    print("\n模块测试结果:")
    validation_passed = sum(1 for f, _, s in results if 'validation/' in f and s)
    validation_total = len(validation_tests)
    print(f"验证模块: {validation_passed}/{validation_total}")

    merge_passed = sum(1 for f, _, s in results if 'merge/' in f and s)
    merge_total = len(merge_tests)
    print(f"合并模块: {merge_passed}/{merge_total}")

    if failed == 0:
        print(f"\n🎉 所有测试通过！")
        sys.exit(0)
    else:
        print(f"\n⚠️  {failed} 个测试失败，请检查输出")
        sys.exit(1)


if __name__ == "__main__":
    main()
