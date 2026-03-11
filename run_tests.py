#!/usr/bin/env python3
"""
快速运行所有测试的脚本
"""

import subprocess
import sys
import os

# 添加父目录到Python路径，以便导入excel_validator模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def run_test(test_file, description):
    """运行单个测试"""
    print(f"\n{'='*60}")
    print(f"运行测试: {description}")
    print(f"{'='*60}")
    try:
        # 设置环境变量，包含父目录到Python路径
        env = os.environ.copy()
        project_root = os.path.dirname(os.path.abspath(__file__))
        if 'PYTHONPATH' in env:
            env['PYTHONPATH'] = project_root + os.pathsep + env['PYTHONPATH']
        else:
            env['PYTHONPATH'] = project_root

        result = subprocess.run(
            [sys.executable, f"tests/{test_file}"],
            capture_output=True,
            text=True,
            timeout=60,
            env=env
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
    tests = [
        ('test_standard.py', '标准测试用例（基于test_abnormal_detail.py）'),
        ('test_abnormal_detail.py', '异常详情中的字符串格式测试'),
        ('test_direct.py', '直接测试string_columns效果'),
        ('test_final.py', '最终验证string_columns的有效范围'),
        ('test_string_columns.py', 'string_columns参数测试'),
        ('test_realistic.py', '现实场景测试'),
    ]

    print("=== Excel 数据验证工具 - 测试套件 ===")
    print(f"将运行 {len(tests)} 个测试")

    results = []
    for test_file, description in tests:
        success = run_test(test_file, description)
        results.append((test_file, description, success))

    # 汇总结果
    print(f"\n{'='*60}")
    print("测试结果汇总")
    print(f"{'='*60}")

    passed = sum(1 for _, _, success in results if success)
    failed = sum(1 for _, _, success in results if not success)

    print(f"通过: {passed}/{len(tests)}")
    print(f"失败: {failed}/{len(tests)}")

    print("\n详细结果:")
    for test_file, description, success in results:
        status = "✅ 通过" if success else "❌ 失败"
        print(f"  {status} - {description} ({test_file})")

    if failed == 0:
        print(f"\n🎉 所有测试通过！")
        sys.exit(0)
    else:
        print(f"\n⚠️  {failed} 个测试失败，请检查输出")
        sys.exit(1)

if __name__ == "__main__":
    main()