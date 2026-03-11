"""
Excel工具主程序
提供统一的数据验证和表合并接口
"""

import argparse
import sys
import os
from typing import Dict

# 添加父目录到路径，以便导入modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.validation import process_excel_with_validation
from modules.merge import merge_excel_tables


def run_validation(args):
    """运行数据验证功能"""
    print("=== 数据验证模式 ===\n")

    try:
        result = process_excel_with_validation(
            input_file=args.input,
            sheet_name=args.sheet,
            group_columns=args.group_columns.split(','),
            compare_columns=args.compare_columns.split(','),
            output_columns=args.output_columns.split(',') if args.output_columns else None,
            output_file=args.output,
            string_columns=args.string_columns.split(',') if args.string_columns else None,
            abnormal_detail_columns=args.abnormal_detail_columns.split(',') if args.abnormal_detail_columns else None
        )

        print("\n验证完成！")
        print(f"输出文件: {args.output}")
        print("\n分组结果:")
        print(result)

    except Exception as e:
        print(f"\n❌ 验证失败: {e}")
        sys.exit(1)


def parse_match_columns(columns_input: str) -> Dict[str, str]:
    """
    解析匹配列参数，支持多种格式

    Args:
        columns_input: 用户输入的列名

    Returns:
        Dict[str, str]: 表A列名到表B列名的映射

    Examples:
        "ID" → {"ID": "ID"}
        "ID:员工编号" → {"ID": "员工编号"}
        "ID:员工编号,部门:部门编码" → {"ID": "员工编号", "部门": "部门编码"}
    """
    mapping = {}

    # 处理空输入
    if not columns_input.strip():
        return mapping

    if ':' in columns_input:
        # 新格式：A:B,C:D
        pairs = columns_input.split(',')
        for pair in pairs:
            if ':' in pair:
                a_col, b_col = pair.split(':', 1)
                mapping[a_col.strip()] = b_col.strip()
            else:
                # 如果只有一侧有冒号，默认为相同列名
                mapping[pair.strip()] = pair.strip()
    else:
        # 旧格式：单个列或多列逗号分隔
        columns = [col.strip() for col in columns_input.split(',')]
        for col in columns:
            if col:  # 跳过空字符串
                mapping[col] = col

    return mapping


def run_merge(args):
    """运行表合并功能"""
    print("=== 表合并模式 ===\n")

    try:
        # 解析匹配列参数
        match_columns = parse_match_columns(args.match_columns)

        result = merge_excel_tables(
            table_a_file=args.table_a,
            table_a_sheet=args.table_a_sheet,
            table_b_file=args.table_b,
            table_b_sheet=args.table_b_sheet,
            match_columns=match_columns,
            table_a_extra_columns=args.table_a_extra_columns.split(',') if args.table_a_extra_columns else None,
            table_b_extra_columns=args.table_b_extra_columns.split(',') if args.table_b_extra_columns else None,
            output_file=args.output,
            string_columns=args.string_columns.split(',') if args.string_columns else None
        )

        print("\n合并完成！")
        print(f"输出文件: {args.output}")
        print("\n合并结果:")
        print(result)

    except Exception as e:
        print(f"\n❌ 合并失败: {e}")
        sys.exit(1)


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='Excel数据验证和表合并工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  数据验证:
    python main.py validate -i data.xlsx -s Sheet1 -g 部门 -c 计划值,实际值 -o result.xlsx

  表合并（相同列名）:
    python main.py merge -a table_a.xlsx -A Sheet1 -b table_b.xlsx -B Sheet1 -m ID -a_extra 姓名,部门 -b_extra 职位,薪资 -o merged.xlsx

  表合并（不同列名）:
    python main.py merge -a table_a.xlsx -A Sheet1 -b table_b.xlsx -B Sheet1 -m "ID:员工编号,部门:部门编码" -a_extra 姓名,职位 -b_extra 薪资,入职日期 -o merged.xlsx
        '''
    )

    # 子命令
    subparsers = parser.add_subparsers(dest='command', help='操作命令')

    # 验证命令
    validate_parser = subparsers.add_parser('validate', help='数据验证')
    validate_parser.add_argument('-i', '--input', required=True, help='输入Excel文件路径')
    validate_parser.add_argument('-s', '--sheet', required=True, help='工作表名称')
    validate_parser.add_argument('-g', '--group-columns', required=True, help='分组列名（逗号分隔）')
    validate_parser.add_argument('-c', '--compare-columns', required=True, help='比较列名（逗号分隔，2列）')
    validate_parser.add_argument('-o', '--output', default='validation_result.xlsx', help='输出文件名')
    validate_parser.add_argument('--output-columns', help='输出列名（逗号分隔）')
    validate_parser.add_argument('--string-columns', help='字符串列名（逗号分隔）')
    validate_parser.add_argument('--abnormal-detail-columns', help='异常详情列名（逗号分隔）')

    # 合并命令
    merge_parser = subparsers.add_parser('merge', help='表合并')
    merge_parser.add_argument('-a', '--table-a', required=True, help='表A的Excel文件路径')
    merge_parser.add_argument('-A', '--table-a-sheet', required=True, help='表A的工作表名称')
    merge_parser.add_argument('-b', '--table-b', required=True, help='表B的Excel文件路径')
    merge_parser.add_argument('-B', '--table-b-sheet', required=True, help='表B的工作表名称')
    merge_parser.add_argument('-m', '--match-columns', required=True, help='匹配列名。支持格式："A:B,C:D"（表A列名:表B列名）或 "A,B"（自动映射为 A:A,B:B）')
    merge_parser.add_argument('--table-a-extra-columns', help='表A额外列名（逗号分隔）')
    merge_parser.add_argument('--table-b-extra-columns', help='表B额外列名（逗号分隔）')
    merge_parser.add_argument('-o', '--output', default='merge_result.xlsx', help='输出文件名')
    merge_parser.add_argument('--string-columns', help='字符串列名（逗号分隔）')

    # 解析参数
    args = parser.parse_args()

    # 执行对应的命令
    if args.command == 'validate':
        run_validation(args)
    elif args.command == 'merge':
        run_merge(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
