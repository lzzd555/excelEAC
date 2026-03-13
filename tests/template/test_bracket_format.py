"""
测试 Excel 外部引用索引格式（如 [3]SheetName!A:R）
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import re

TEST_DIR = os.path.dirname(os.path.abspath(__file__))


def test_bracket_regex():
    """测试正则表达式匹配"""
    print("=" * 70)
    print("测试：正则表达式匹配 Excel 外部引用索引格式")
    print("=" * 70 + "\n")

    # 新的正则表达式
    bracket_pattern = r"(\[[^\]]+\][^!'\s]+)!([A-Z]+\d*(?::[A-Z]*\d*)?)"

    test_cases = [
        ("[3]合同汇总分析表!A:R", True, "[3]合同汇总分析表", "A:R"),
        ("[4]CPQ标识!G:H", True, "[4]CPQ标识", "G:H"),
        ("[6]ESDP-Bpart!A:A", True, "[6]ESDP-Bpart", "A:A"),
        ("'[6]ESDP-Bpart'!A:A", False, None, None),  # 带单引号的由 quoted_pattern 处理
        ("Sheet1!A1", False, None, None),  # 无方括号的由 unquoted_pattern 处理
        ("=VLOOKUP(E2,[3]合同汇总分析表!A:R,2,0)", True, "[3]合同汇总分析表", "A:R"),
    ]

    all_passed = True
    for formula, should_match, expected_ref, expected_cell in test_cases:
        match = re.search(bracket_pattern, formula, re.IGNORECASE)
        if should_match:
            if match:
                ref = match.group(1)
                cell = match.group(2)
                if ref == expected_ref and cell == expected_cell:
                    print(f"  ✅ '{formula}'")
                    print(f"     → 匹配: '{ref}'!'{cell}'")
                else:
                    print(f"  ❌ '{formula}'")
                    print(f"     期望: '{expected_ref}'!'{expected_cell}', 实际: '{ref}'!'{cell}'")
                    all_passed = False
            else:
                print(f"  ❌ '{formula}' → 未匹配（应该匹配）")
                all_passed = False
        else:
            if match:
                print(f"  ❌ '{formula}' → 匹配了（不应该匹配）: '{match.group(1)}'")
                all_passed = False
            else:
                print(f"  ✅ '{formula}' → 未匹配（正确，由其他正则处理）")

    return all_passed


def test_full_replacement():
    """测试完整的公式替换流程"""
    print("\n" + "=" * 70)
    print("测试：完整公式替换流程")
    print("=" * 70 + "\n")

    from modules.template_generator import replace_sheet_references

    alias_to_info = {
        '合同汇总分析表': {
            'file_path': '/path/to/合同汇总分析表.xlsx',
            'sheet_name': '合同汇总分析表'
        },
        'cpq标识': {
            'file_path': '/path/to/CPQ标识.xlsx',
            'sheet_name': 'CPQ标识'
        },
        'esdp-bpart': {
            'file_path': '/path/to/ESDP_Bpart.xlsx',
            'sheet_name': 'ESDP-Bpart'
        }
    }

    test_cases = [
        ("=VLOOKUP(E2,[3]合同汇总分析表!A:R,2,0)", "[合同汇总分析表.xlsx]合同汇总分析表"),
        ("=VLOOKUP(E2,[4]CPQ标识!G:H,2,0)", "[CPQ标识.xlsx]CPQ标识"),
        ("=COUNTIFS([6]ESDP-Bpart!A:A,A2)", "[ESDP_Bpart.xlsx]ESDP-Bpart"),
    ]

    all_passed = True
    for formula, expected in test_cases:
        result = replace_sheet_references(formula, alias_to_info, row_offset=0)
        if expected in result:
            print(f"  ✅ 原始: {formula}")
            print(f"     结果: {result}")
        else:
            print(f"  ❌ 原始: {formula}")
            print(f"     结果: {result}")
            print(f"     期望包含: {expected}")
            all_passed = False

    return all_passed


if __name__ == "__main__":
    print("=" * 70)
    print("Excel 外部引用索引格式测试")
    print("=" * 70 + "\n")

    regex_passed = test_bracket_regex()
    replacement_passed = test_full_replacement()

    print("\n" + "=" * 70)
    print("总结")
    print("=" * 70)
    print(f"  正则表达式测试: {'✅ 通过' if regex_passed else '❌ 失败'}")
    print(f"  完整替换测试: {'✅ 通过' if replacement_passed else '❌ 失败'}")
    print("=" * 70)
