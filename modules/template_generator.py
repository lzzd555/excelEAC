"""
Excel模板生成器模块
基于模板生成Excel，支持多数据源、公式保留和列映射
"""

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
import os
import re


@dataclass
class DataColumnMapping:
    """列映射配置"""
    source_column: str  # 源列名
    target_column: str  # 目标列名（与模板列名匹配）


@dataclass
class DataSourceConfig:
    """数据源配置"""
    file_path: str                           # 数据文件路径
    sheet_name: str                          # sheet名称
    column_mappings: List[DataColumnMapping] # 列映射集合
    alias: str = ""                          # 别名（如"sheet0", "sheet1"）


def read_template_structure(
    template_file: str,
    template_sheet: str
) -> Tuple[List[str], Dict[str, str]]:
    """
    读取模板结构，包括列名、顺序和公式

    Args:
        template_file: 模板文件路径
        template_sheet: 模板sheet名称

    Returns:
        Tuple[List[str], Dict[str, str]]: (列名列表, 列名到公式模板的映射)
    """
    print(f"正在读取模板: {template_file} 的 {template_sheet} 工作表...")

    # 使用openpyxl读取以保留公式
    wb = openpyxl.load_workbook(template_file, data_only=False)

    if template_sheet not in wb.sheetnames:
        raise ValueError(f"模板sheet '{template_sheet}' 不存在。可用的sheet: {wb.sheetnames}")

    ws = wb[template_sheet]

    # 读取第一行作为列名
    column_names = []
    for cell in ws[1]:
        if cell.value is not None:
            column_names.append(str(cell.value))
        else:
            break

    # 读取第二行的公式（如果存在）
    formula_templates = {}
    if ws.max_row >= 2:
        for col_idx, col_name in enumerate(column_names, start=1):
            cell = ws.cell(row=2, column=col_idx)
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                formula_templates[col_name] = cell.value

    wb.close()

    print(f"模板列名: {column_names}")
    print(f"公式列: {list(formula_templates.keys())}")

    return column_names, formula_templates


def read_data_source(
    config: DataSourceConfig,
    string_columns: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    读取单个数据源并应用列映射

    Args:
        config: 数据源配置
        string_columns: 需要保持为字符串的列名列表

    Returns:
        pd.DataFrame: 处理后的数据（列名为目标列名）
    """
    print(f"正在读取数据源: {config.file_path} 的 {config.sheet_name} 工作表...")

    # 构建dtype字典
    dtype_dict = {}
    if string_columns:
        for mapping in config.column_mappings:
            if mapping.target_column in string_columns:
                dtype_dict[mapping.source_column] = 'string'

    df = pd.read_excel(config.file_path, sheet_name=config.sheet_name, dtype=dtype_dict)

    # 验证源列是否存在
    source_columns = [m.source_column for m in config.column_mappings]
    missing_cols = [col for col in source_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"数据源 {config.file_path} 中缺少列: {missing_cols}")

    # 应用列映射：重命名列
    rename_dict = {m.source_column: m.target_column for m in config.column_mappings}
    df = df.rename(columns=rename_dict)

    # 只保留映射后的目标列
    target_columns = [m.target_column for m in config.column_mappings]
    df = df[[col for col in target_columns if col in df.columns]]

    print(f"数据源 '{config.alias}' 数据: {len(df)} 行, {len(df.columns)} 列")

    return df


def merge_data_by_row(
    data_sources: List[Tuple[str, pd.DataFrame]],
    template_columns: List[str]
) -> pd.DataFrame:
    """
    按行号对齐合并多个数据源

    Args:
        data_sources: 数据源列表 [(alias, DataFrame), ...]
        template_columns: 模板列名列表（用于确定输出列顺序）

    Returns:
        pd.DataFrame: 合并后的数据
        - 行数 = max(各数据源行数)
        - 第i行 = 从各数据源取第i行（按列映射）
        - 如果多表映射到同一列，验证值是否一致
        - 模板中存在但数据源中没有的列会被创建为空列
    """
    if not data_sources:
        return pd.DataFrame(columns=template_columns)

    # 计算最大行数
    max_rows = max(len(df) for _, df in data_sources) if data_sources else 0
    print(f"合并数据: 最大行数 = {max_rows}")

    # 收集所有数据源的列
    all_data_columns = set()
    for _, df in data_sources:
        all_data_columns.update(df.columns)

    # 构建合并后的数据
    merged_data = []

    for row_idx in range(max_rows):
        new_row = {}

        # 从每个数据源取第row_idx行
        for alias, df in data_sources:
            if row_idx < len(df):
                row_data = df.iloc[row_idx]
                for col in df.columns:
                    value = row_data[col]

                    # 如果多表映射到同一列，验证值是否一致
                    if col in new_row:
                        existing_value = new_row[col]
                        # 如果两个值都非空且不相等，发出警告
                        if pd.notna(existing_value) and pd.notna(value):
                            if existing_value != value:
                                print(f"警告: 行{row_idx + 1} 列'{col}' 值不一致: "
                                      f"'{existing_value}' vs '{value}' (来自 {alias})")
                    else:
                        new_row[col] = value

        merged_data.append(new_row)

    merged_df = pd.DataFrame(merged_data)

    # 确保所有模板列都存在（包括公式列等数据源中没有的列）
    for col in template_columns:
        if col not in merged_df.columns:
            merged_df[col] = None

    # 确保列顺序与模板一致
    final_columns = [col for col in template_columns if col in merged_df.columns]
    # 添加模板中不存在但数据中有的列
    extra_columns = [col for col in merged_df.columns if col not in final_columns]
    final_columns.extend(extra_columns)

    merged_df = merged_df[final_columns]

    print(f"合并完成: {len(merged_df)} 行, {len(merged_df.columns)} 列")

    return merged_df


def parse_formula_references(formula: str) -> List[Tuple[str, str, str]]:
    """
    解析公式中的sheet引用

    Args:
        formula: Excel公式字符串

    Returns:
        List[Tuple[str, str, str]]: [(sheet别名, 列字母, 行号), ...]
    """
    # 匹配模式: sheet0!A1, sheet1!B2 等
    pattern = r"(sheet\d+)!([A-Z]+)(\d+)"
    matches = re.findall(pattern, formula, re.IGNORECASE)

    return [(m[0].lower(), m[1].upper(), int(m[2])) for m in matches]


def replace_sheet_references(
    formula: str,
    alias_to_info: Dict[str, Dict[str, str]],
    row_offset: int = 0
) -> str:
    """
    替换公式中的sheet引用为外部文件引用，并调整行号

    Args:
        formula: 原始公式字符串
        alias_to_info: 别名到文件信息的映射 {alias: {'file_path': str, 'sheet_name': str}}
        row_offset: 行号偏移量（用于调整输出行号）

    Returns:
        str: 替换后的公式
        格式: sheet0!A1 -> '[filename.xlsx]actual_sheet_name'!A1
             sheet0!A:A -> '[filename.xlsx]actual_sheet_name'!A:A (整列引用)
    """
    # 匹配sheet引用，支持多种格式：
    # - sheet0!A1 (单元格)
    # - sheet0!A:A (整列)
    # - sheet0!A1:A10 (列范围)
    # - sheet0!A1:B10 (单元格范围)
    sheet_pattern = r"(sheet\d+)!([A-Z]+\d*(?::[A-Z]*\d*)?)"

    def replace_sheet_match(match):
        alias = match.group(1).lower()
        cell_ref = match.group(2).upper()

        if alias in alias_to_info:
            info = alias_to_info[alias]
            file_path = info['file_path']
            actual_sheet_name = info['sheet_name']
            file_name = os.path.basename(file_path)

            # 处理单元格引用，调整行号
            # 检查是否是整列引用 (如 A:A) 或范围引用 (如 A1:B10)
            if ':' in cell_ref:
                # 范围引用，需要处理两部分
                parts = cell_ref.split(':')
                adjusted_parts = []

                for part in parts:
                    # 提取列字母和行号
                    col_match = re.match(r'([A-Z]+)(\d*)', part)
                    if col_match:
                        col = col_match.group(1)
                        row_str = col_match.group(2)
                        if row_str:
                            # 有行号，需要调整
                            row = int(row_str) + row_offset
                            adjusted_parts.append(f"{col}{row}")
                        else:
                            # 只有列字母（整列引用），不调整
                            adjusted_parts.append(col)
                    else:
                        adjusted_parts.append(part)

                adjusted_ref = ':'.join(adjusted_parts)
            else:
                # 单个单元格引用
                col_match = re.match(r'([A-Z]+)(\d+)', cell_ref)
                if col_match:
                    col = col_match.group(1)
                    row = int(col_match.group(2)) + row_offset
                    adjusted_ref = f"{col}{row}"
                else:
                    # 只有列字母，不调整
                    adjusted_ref = cell_ref

            return f"'[{file_name}]{actual_sheet_name}'!{adjusted_ref}"
        else:
            return match.group(0)

    result = re.sub(sheet_pattern, replace_sheet_match, formula, flags=re.IGNORECASE)

    # 然后调整本地单元格引用的行号（不包含sheet引用的）
    def adjust_local_ref(match):
        col = match.group(1)
        row = int(match.group(2)) + row_offset
        return f"{col}{row}"

    # 匹配本地单元格引用（不在单引号内，不跟在sheet后面的）
    local_pattern = r"(?<![A-Za-z!'\"\\])([A-Z]+)(\d+)(?![A-Za-z])"
    result = re.sub(local_pattern, adjust_local_ref, result)

    return result


def apply_formulas_to_output(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    formula_columns: List[str],
    formula_templates: Dict[str, str],
    alias_to_info: Dict[str, Dict[str, str]],
    start_row: int = 2
) -> None:
    """
    将公式应用到输出文件

    Args:
        ws: openpyxl工作表对象
        formula_columns: 公式列名列表
        formula_templates: 列名到公式模板的映射
        alias_to_info: 别名到文件信息的映射 {alias: {'file_path': str, 'sheet_name': str}}
        start_row: 开始应用公式的行号（默认为2，跳过标题行）
    """
    # 获取列名到列号的映射
    header_row = 1
    col_name_to_idx = {}
    for col_idx, cell in enumerate(ws[header_row], start=1):
        if cell.value:
            col_name_to_idx[str(cell.value)] = col_idx

    # 应用公式
    for col_name in formula_columns:
        if col_name not in col_name_to_idx:
            print(f"警告: 公式列 '{col_name}' 不在输出列中")
            continue

        if col_name not in formula_templates:
            print(f"警告: 没有找到列 '{col_name}' 的公式模板")
            continue

        col_idx = col_name_to_idx[col_name]
        formula_template = formula_templates[col_name]

        # 为每一行应用公式
        for row_idx in range(start_row, ws.max_row + 1):
            # 计算行号偏移（相对于模板中的行号）
            row_offset = row_idx - start_row

            # 替换公式中的引用
            new_formula = replace_sheet_references(formula_template, alias_to_info, row_offset)

            # 写入公式
            ws.cell(row=row_idx, column=col_idx).value = new_formula

    print(f"已应用公式到列: {formula_columns}")


def generate_excel_from_template(
    template_file: str,
    template_sheet: str,
    formula_columns: List[str],
    data_sources: List[Dict],
    output_file: str,
    string_columns: Optional[List[str]] = None,
    use_external_refs: bool = True
) -> pd.DataFrame:
    """
    基于模板生成Excel文件

    Args:
        template_file: 模板文件路径
        template_sheet: 模板sheet名称
        formula_columns: 公式列集合（这些列需要保留公式并处理sheet引用）
        data_sources: 数据源集合，每个元素为字典格式:
            {
                'file_path': str,       # 数据文件路径
                'sheet_name': str,      # sheet名称
                'column_mappings': List[Dict],  # 列映射集合 [{'source': str, 'target': str}, ...]
                'alias': str            # 别名（如"sheet0", "sheet1"）
            }
        output_file: 输出文件路径
        string_columns: 字符串列列表（保持前导零等格式）
        use_external_refs: 是否使用外部文件引用公式。
            True: 公式使用外部引用（如 ='[sales.xlsx]Sheet1'!A1），需Excel打开
            False: 直接写入数据值，只有本地计算公式（如 =A1+B1）保留公式

    Returns:
        pd.DataFrame: 生成的数据

    示例:
        data_sources = [
            {
                'file_path': 'sales.xlsx',
                'sheet_name': 'Sheet1',
                'column_mappings': [
                    {'source': 'SalesAmt', 'target': 'Sales'},
                    {'source': 'Date', 'target': 'Date'}
                ],
                'alias': 'sheet0'
            },
            {
                'file_path': 'costs.xlsx',
                'sheet_name': 'Sheet1',
                'column_mappings': [
                    {'source': 'CostAmt', 'target': 'Cost'},
                    {'source': 'Date', 'target': 'Date'}
                ],
                'alias': 'sheet1'
            }
        ]
    """
    # 确保输出文件路径是绝对路径
    if not os.path.isabs(output_file):
        output_file = os.path.join(os.getcwd(), output_file)

    print("=== Excel模板生成器 ===\n")

    # 1. 验证阶段：检查所有文件和sheet是否存在
    print("1. 验证文件...")

    if not os.path.exists(template_file):
        raise FileNotFoundError(f"模板文件不存在: {template_file}")

    for ds in data_sources:
        if not os.path.exists(ds['file_path']):
            raise FileNotFoundError(f"数据文件不存在: {ds['file_path']}")

    print("   所有文件验证通过\n")

    # 2. 模板分析：读取模板列名、顺序和公式模式
    print("2. 分析模板...")
    template_columns, formula_templates = read_template_structure(template_file, template_sheet)

    # 收集模板中的公式（用于后续应用）
    template_formulas = {}
    for col_name in formula_columns:
        if col_name in formula_templates:
            template_formulas[col_name] = formula_templates[col_name]
        else:
            print(f"   警告: 公式列 '{col_name}' 在模板中没有公式")

    print()

    # 3. 数据加载：读取各数据源，应用列映射
    print("3. 加载数据源...")

    # 构建数据源配置
    loaded_data_sources = []
    alias_to_info = {}

    for ds in data_sources:
        # 构建列映射
        column_mappings = [
            DataColumnMapping(source_column=m['source'], target_column=m['target'])
            for m in ds['column_mappings']
        ]

        config = DataSourceConfig(
            file_path=ds['file_path'],
            sheet_name=ds['sheet_name'],
            column_mappings=column_mappings,
            alias=ds.get('alias', '')
        )

        # 读取数据
        df = read_data_source(config, string_columns)
        loaded_data_sources.append((config.alias, df))

        # 记录别名到文件信息和sheet名称的映射
        alias_to_info[config.alias.lower()] = {
            'file_path': config.file_path,
            'sheet_name': config.sheet_name
        }

    print()

    # 4. 数据合并：按行号对齐合并
    print("4. 合并数据...")
    merged_df = merge_data_by_row(loaded_data_sources, template_columns)

    print()

    # 5. 输出生成：写入数据，应用公式，保存文件
    print("5. 生成输出文件...")

    if use_external_refs:
        print("   模式: 外部引用公式（需Excel打开）")
    else:
        print("   模式: 直接写入数据值")

    # 创建输出文件
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        output_df = merged_df.copy()

        if use_external_refs:
            # 外部引用模式：清空公式列的数据，后面填充外部引用公式
            for col in formula_columns:
                if col in output_df.columns:
                    output_df[col] = None
        else:
            # 直接数据模式：保留数据值，只处理本地计算公式
            # 找出哪些公式列包含外部引用（sheet0!, sheet1!等）
            external_ref_columns = []
            local_formula_columns = []

            for col in formula_columns:
                if col in template_formulas:
                    formula = template_formulas[col]
                    # 检查是否包含外部引用
                    if re.search(r'sheet\d+!', formula, re.IGNORECASE):
                        external_ref_columns.append(col)
                    else:
                        local_formula_columns.append(col)

            # 清空本地计算公式列（这些会填充公式）
            for col in local_formula_columns:
                if col in output_df.columns:
                    output_df[col] = None

            # 外部引用列保留数据值
            print(f"   数据值列: {external_ref_columns}")
            print(f"   本地公式列: {local_formula_columns}")

        output_df.to_excel(writer, sheet_name='结果', index=False)

        # 获取工作表
        ws = writer.sheets['结果']

        # 应用公式
        if formula_columns and template_formulas:
            if use_external_refs:
                # 外部引用模式：应用所有公式
                apply_formulas_to_output(
                    ws,
                    formula_columns,
                    template_formulas,
                    alias_to_info,
                    start_row=2
                )
            else:
                # 直接数据模式：只应用本地计算公式
                for col in formula_columns:
                    if col in template_formulas:
                        formula = template_formulas[col]
                        # 只处理不包含外部引用的公式
                        if not re.search(r'sheet\d+!', formula, re.IGNORECASE):
                            # 获取列号
                            col_idx = None
                            for c_idx, c_name in enumerate(output_df.columns, start=1):
                                if c_name == col:
                                    col_idx = c_idx
                                    break

                            if col_idx:
                                # 为每一行应用公式
                                for row_idx in range(2, ws.max_row + 1):
                                    row_offset = row_idx - 2
                                    # 只调整本地引用的行号
                                    adjusted_formula = replace_sheet_references(
                                        formula, {}, row_offset
                                    )
                                    ws.cell(row=row_idx, column=col_idx).value = adjusted_formula

                                print(f"   已应用公式: {col} = {formula}")

        # 处理字符串列格式
        if string_columns:
            from openpyxl.styles import Font

            for col_idx, col_name in enumerate(output_df.columns, start=1):
                if col_name in string_columns:
                    # 设置列宽
                    ws.column_dimensions[get_column_letter(col_idx)].width = 15

                    # 设置格式为文本
                    for row in range(2, len(output_df) + 2):
                        cell = ws.cell(row=row, column=col_idx)
                        cell.number_format = '@'

                        # 确保数据是字符串
                        if pd.notna(output_df.iloc[row - 2][col_name]):
                            cell.value = str(output_df.iloc[row - 2][col_name])

    print(f"\n输出文件已保存: {output_file}")

    return merged_df


# 命令行解析辅助函数
def parse_column_mappings(mappings_str: str) -> List[Dict[str, str]]:
    """
    解析列映射字符串

    Args:
        mappings_str: 列映射字符串，格式: "SourceCol1:TargetCol1,SourceCol2:TargetCol2"

    Returns:
        List[Dict[str, str]]: 列映射列表
    """
    mappings = []
    pairs = mappings_str.split(',')

    for pair in pairs:
        if ':' in pair:
            source, target = pair.split(':', 1)
            mappings.append({
                'source': source.strip(),
                'target': target.strip()
            })
        else:
            # 如果没有冒号，源列和目标列相同
            col = pair.strip()
            mappings.append({
                'source': col,
                'target': col
            })

    return mappings


if __name__ == "__main__":
    # 示例使用
    print("=== Excel模板生成器示例 ===\n")

    # 注意：此示例需要先创建测试文件
    # 实际使用时请确保文件存在

    print("使用方法:")
    print("""
    from modules.template_generator import generate_excel_from_template

    data_sources = [
        {
            'file_path': 'sales.xlsx',
            'sheet_name': 'Sheet1',
            'column_mappings': [
                {'source': 'SalesAmt', 'target': 'Sales'},
                {'source': 'Date', 'target': 'Date'}
            ],
            'alias': 'sheet0'
        },
        {
            'file_path': 'costs.xlsx',
            'sheet_name': 'Sheet1',
            'column_mappings': [
                {'source': 'CostAmt', 'target': 'Cost'},
                {'source': 'Date', 'target': 'Date'}
            ],
            'alias': 'sheet1'
        }
    ]

    result = generate_excel_from_template(
        template_file='template.xlsx',
        template_sheet='Sheet1',
        formula_columns=['Total', 'Profit'],
        data_sources=data_sources,
        output_file='output.xlsx',
        string_columns=['Date']
    )
    """)
