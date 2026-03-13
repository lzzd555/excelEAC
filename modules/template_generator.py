"""
Excel模板生成器模块
基于模板生成Excel，支持多数据源、公式保留和列映射
"""

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Fill, Border, Alignment, Protection, PatternFill
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


@dataclass
class CellStyle:
    """单元格样式"""
    font: Optional[Dict[str, Any]] = None
    fill: Optional[Dict[str, Any]] = None
    border: Optional[Dict[str, Any]] = None
    alignment: Optional[Dict[str, Any]] = None
    number_format: Optional[str] = None
    protection: Optional[Dict[str, Any]] = None


def copy_cell_style(source_cell, target_cell) -> None:
    """
    复制单元格样式

    Args:
        source_cell: 源单元格
        target_cell: 目标单元格
    """
    if source_cell.has_style:
        # 复制字体
        if source_cell.font:
            target_cell.font = Font(
                name=source_cell.font.name,
                size=source_cell.font.size,
                bold=source_cell.font.bold,
                italic=source_cell.font.italic,
                vertAlign=source_cell.font.vertAlign,
                underline=source_cell.font.underline,
                strike=source_cell.font.strike,
                color=source_cell.font.color
            )

        # 复制填充
        if source_cell.fill and source_cell.fill.fill_type:
            fill_type = source_cell.fill.fill_type
            fg_color = source_cell.fill.fgColor
            bg_color = source_cell.fill.bgColor
            if fill_type and fg_color:
                target_cell.fill = PatternFill(
                    fill_type=fill_type,
                    start_color=fg_color.rgb if fg_color.rgb else fg_color.idx,
                    end_color=bg_color.rgb if bg_color else None
                )

        # 复制边框
        if source_cell.border:
            target_cell.border = Border(
                left=source_cell.border.left,
                right=source_cell.border.right,
                top=source_cell.border.top,
                bottom=source_cell.border.bottom,
                diagonal=source_cell.border.diagonal,
                diagonal_direction=source_cell.border.diagonal_direction,
                outline=source_cell.border.outline,
                horizontal=source_cell.border.horizontal,
                vertical=source_cell.border.vertical
            )

        # 复制对齐
        if source_cell.alignment:
            target_cell.alignment = Alignment(
                horizontal=source_cell.alignment.horizontal,
                vertical=source_cell.alignment.vertical,
                text_rotation=source_cell.alignment.text_rotation,
                wrap_text=source_cell.alignment.wrap_text,
                shrink_to_fit=source_cell.alignment.shrink_to_fit,
                indent=source_cell.alignment.indent
            )

        # 复制数字格式
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format

        # 复制保护
        if source_cell.protection:
            target_cell.protection = Protection(
                locked=source_cell.protection.locked,
                hidden=source_cell.protection.hidden
            )


def read_external_links(xlsx_file: str) -> Dict[int, str]:
    """
    读取Excel文件中的外部链接映射

    Excel会将外部引用的文件名替换为数字索引（如[3]SheetName）
    这个函数从xlsx文件的内部XML结构中读取索引与文件名的映射关系

    Args:
        xlsx_file: Excel文件路径

    Returns:
        Dict[int, str]: {索引号: 文件名} 的映射字典
    """
    import zipfile
    import xml.etree.ElementTree as ET

    link_mapping = {}

    try:
        with zipfile.ZipFile(xlsx_file, 'r') as z:
            # 检查是否有外部链接目录
            external_links_dir = 'xl/externalLinks/'
            link_files = [name for name in z.namelist() if name.startswith(external_links_dir) and name.endswith('.xml')]

            for link_file in link_files:
                try:
                    content = z.read(link_file).decode('utf-8')
                    root = ET.fromstring(content)

                    # 提取索引号（从文件名 externalLink1.xml 中提取数字）
                    import re as re_module
                    match = re_module.search(r'externalLink(\d+)\.xml', link_file)
                    if match:
                        index = int(match.group(1))

                        # 查找外部链接的文件路径
                        # 命名空间可能是 http://schemas.openxmlformats.org/officeDocument/2006/relationships
                        for elem in root.iter():
                            # 查找包含 target 或 Target 属性的元素
                            if 'Target' in elem.attrib:
                                target = elem.attrib['Target']
                                # 提取文件名（去掉路径）
                                filename = os.path.basename(target)
                                link_mapping[index] = filename
                                break
                            elif 'target' in elem.attrib:
                                target = elem.attrib['target']
                                filename = os.path.basename(target)
                                link_mapping[index] = filename
                                break
                except Exception as e:
                    print(f"   警告: 读取外部链接文件 {link_file} 失败: {e}")
                    continue

    except Exception as e:
        print(f"   警告: 读取外部链接失败: {e}")

    if link_mapping:
        print(f"   外部链接映射: {link_mapping}")

    return link_mapping


def read_template_structure(
    template_file: str,
    template_sheet: str
) -> Tuple[List[str], Dict[str, str], openpyxl.worksheet.worksheet.Worksheet]:
    """
    读取模板结构，包括列名、顺序和公式

    Args:
        template_file: 模板文件路径
        template_sheet: 模板sheet名称

    Returns:
        Tuple[List[str], Dict[str, str], Worksheet]: (列名列表, 列名到公式模板的映射, 模板工作表对象)
    """
    print(f"正在读取模板: {template_file} 的 {template_sheet} 工作表...")

    # 读取外部链接映射（索引号 -> 文件名）
    external_links = read_external_links(template_file)

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
                formula = cell.value

                # 如果有外部链接映射，替换公式中的索引号为实际文件名
                if external_links:
                    formula = replace_link_indices_with_filenames(formula, external_links)

                formula_templates[col_name] = formula

    # 注意：不要关闭 wb，我们需要返回 ws 用于样式复制
    # wb.close()

    print(f"模板列名: {column_names}")
    print(f"公式列: {list(formula_templates.keys())}")

    return column_names, formula_templates, ws


def replace_link_indices_with_filenames(formula: str, link_mapping: Dict[int, str]) -> str:
    """
    将公式中的外部链接索引号替换为实际文件名

    例如: [3]SheetName!A:R -> [filename.xlsx]SheetName!A:R

    Args:
        formula: 原始公式字符串
        link_mapping: {索引号: 文件名} 的映射字典

    Returns:
        str: 替换后的公式
    """
    import re as re_module

    def replace_index(match):
        index = int(match.group(1))
        if index in link_mapping:
            filename = link_mapping[index]
            return f"[{filename}]"
        return match.group(0)  # 如果没有找到映射，保持原样

    # 匹配 [数字] 格式（如 [3], [4], [6]）
    # 但不要匹配已经是文件名的情况（如 [sales.xlsx]）
    pattern = r'\[(\d+)\](?![^\[]*\.xlsx)'

    result = re_module.sub(pattern, replace_index, formula)
    return result


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

    # 如果没有列映射，说明这个数据源只用于公式引用，不需要提取数据列
    if not config.column_mappings:
        print(f"数据源 '{config.alias}' 仅用于公式引用，不提取数据列")
        # 返回一个空的DataFrame，但保持正确的行数
        return pd.DataFrame(index=range(len(df)))

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

    # 确保字符串列保持字符串格式（防止前置零丢失）
    if string_columns:
        for col in string_columns:
            if col in df.columns:
                df[col] = df[col].astype('string')

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
        List[Tuple[str, str, str]]: [(sheet名/别名, 列字母, 行号), ...]
    """
    # 匹配模式: sheet0!A1, Sheet1!B2, 'MySheet'!C3 等
    pattern = r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!([A-Z]+)(\d+)"
    matches = re.findall(pattern, formula, re.IGNORECASE)

    results = []
    for m in matches:
        sheet_name = m[0] if m[0] else m[1]  # 带引号的sheet名或不带引号的
        results.append((sheet_name.lower(), m[2].upper(), int(m[3])))

    return results


def replace_sheet_references(
    formula: str,
    alias_to_info: Dict[str, Dict[str, str]],
    row_offset: int = 0,
    output_file_path: Optional[str] = None,
    external_links: Optional[Dict[int, str]] = None
) -> str:
    """
    替换公式中的sheet引用为外部文件引用，并调整行号

    Args:
        formula: 原始公式字符串
        alias_to_info: sheet名到文件信息的映射 {sheet_name: {'file_path': str, 'sheet_name': str, 'is_template_self_reference': bool}}
                       key可以是别名、实际的sheet名或模板sheet名
        row_offset: 行号偏移量（用于调整输出行号）
        output_file_path: 输出文件的完整路径。如果引用的是输出文件自身，则返回本地引用格式
        external_links: 外部链接映射 {索引号: 文件名}，用于匹配 Excel 索引号到实际数据源

    Returns:
        str: 替换后的公式
        格式: SheetName!A1 -> '[filename.xlsx]SheetName'!A1
             'Sheet-Name'!A:A -> '[filename.xlsx]Sheet-Name'!A:A
             '[old.xlsx]SheetName'!A1 -> '[new.xlsx]SheetName'!A1
             如果引用的是输出文件自身，则返回本地引用: SheetName!A1
    """
    # 匹配sheet引用，支持多种格式：
    # - sheet0!A1 或 Sheet1!A1 (不带引号的sheet名)
    # - 'ESDP-Bpart'!A:A (带单引号的sheet名)
    # - '[sales.xlsx]Sheet1'!A1 (带文件路径的sheet名)
    # - [3]SheetName!A1 (Excel外部引用索引格式)
    # - 单元格引用: A1
    # - 整列引用: A:A
    # - 范围引用: A1:A10, A1:B10

    # 匹配带单引号的sheet名（可能包含文件路径）: 'SheetName'!CellRef 或 '[filename]SheetName'!CellRef
    quoted_pattern = r"'([^']+)'!([A-Z]+\d*(?::[A-Z]*\d*)?)"

    # 匹配带方括号但无单引号的格式: [xxx]SheetName!CellRef (Excel外部引用索引格式)
    bracket_pattern = r"(\[[^\]]+\][^!'\s]+)!([A-Z]+\d*(?::[A-Z]*\d*)?)"

    # 匹配不带单引号和方括号的sheet名: SheetName!CellRef (sheet名由字母、数字、下划线组成)
    unquoted_pattern = r"([A-Za-z_][A-Za-z0-9_]*)!([A-Z]+\d*(?::[A-Z]*\d*)?)"

    def adjust_cell_ref(cell_ref: str) -> str:
        """调整单元格引用的行号"""
        if ':' in cell_ref:
            parts = cell_ref.split(':')
            adjusted_parts = []
            for part in parts:
                col_match = re.match(r'([A-Z]+)(\d*)', part)
                if col_match:
                    col = col_match.group(1)
                    row_str = col_match.group(2)
                    if row_str:
                        row = int(row_str) + row_offset
                        adjusted_parts.append(f"{col}{row}")
                    else:
                        adjusted_parts.append(col)
                else:
                    adjusted_parts.append(part)
            return ':'.join(adjusted_parts)
        else:
            col_match = re.match(r'([A-Z]+)(\d+)', cell_ref)
            if col_match:
                col = col_match.group(1)
                row = int(col_match.group(2)) + row_offset
                return f"{col}{row}"
            return cell_ref

    def extract_sheet_name(full_reference: str) -> str:
        """
        从完整引用中提取实际的sheet名
        例如: '[sales.xlsx]Sheet1' -> 'Sheet1'
              '[3]Sheet1' -> 'Sheet1'  (Excel外部引用索引格式)
              'ESDP-Bpart' -> 'ESDP-Bpart'
        """
        # 检查是否包含方括号格式 [xxx]sheetname
        # 包括: [filename.xlsx], [数字索引]
        bracket_match = re.match(r'\[.+\](.+)', full_reference)
        if bracket_match:
            return bracket_match.group(1)
        return full_reference

    def find_matching_info(sheet_name: str):
        """查找匹配的sheet信息（支持别名、实际sheet名和数字索引）"""
        # 首先提取实际的sheet名（去掉可能存在的文件路径）
        actual_sheet_name = extract_sheet_name(sheet_name)
        actual_sheet_name_lower = actual_sheet_name.lower()

        # 1. 首先尝试精确匹配（不区分大小写）
        for key, info in alias_to_info.items():
            if key.lower() == actual_sheet_name_lower:
                return info

        # 2. 尝试匹配实际的sheet名
        matching_infos = []
        for key, info in alias_to_info.items():
            if info.get('sheet_name', '').lower() == actual_sheet_name_lower:
                matching_infos.append((key, info))

        # 如果有多个匹配（sheet名称相同），尝试通过文件名匹配
        if len(matching_infos) > 1:
            # 提取引用中的文件名（如果有）
            filename_match = re.search(r'\[([^\]]+)\]', sheet_name)
            if filename_match:
                filename = filename_match.group(1).lower()
                # 查找文件名匹配的数据源
                for key, info in matching_infos:
                    if filename in os.path.basename(info['file_path']).lower():
                        return info
            # 如果没有找到文件名匹配，返回第一个匹配
            return matching_infos[0][1]
        elif matching_infos:
            return matching_infos[0][1]
        
        # 提取方括号中的索引号（如 [0], [1], [3] 等）
        bracket_match = re.match(r'\[(\d+)\]', sheet_name)
        if bracket_match:
            index = bracket_match.group(1)

            # 如果有外部链接映射，使用索引号查找对应的数据源
            if external_links and index in external_links:
                external_filename = external_links[index].lower()

                # 查找文件名匹配的数据源
                for key, info in alias_to_info.items():
                    if external_filename in os.path.basename(info['file_path']).lower():
                        return info

                # 如果没有找到匹配的数据源，尝试使用索引号作为 data_sources 的索引
                if index in alias_to_info:
                    return alias_to_info[index]
            elif index in alias_to_info:
                # 直接用索引号查找（作为 data_sources 的索引）
                return alias_to_info[index]

        return None

    def replace_quoted_match(match):
        full_reference = match.group(1)  # 可能是 'SheetName' 或 '[filename]SheetName'
        cell_ref = match.group(2).upper()

        info = find_matching_info(full_reference)
        if info:
            file_path = info['file_path']
            actual_sheet_name = info['sheet_name']
            adjusted_ref = adjust_cell_ref(cell_ref)

            # 检查是否是本地引用（输出文件自身或模板自引用）
            is_local = (
                (output_file_path and os.path.normpath(file_path) == os.path.normpath(output_file_path)) or
                info.get('is_template_self_reference', False)
            )

            if is_local:
                # 返回本地引用格式
                return f"'{actual_sheet_name}'!{adjusted_ref}"
            else:
                # 返回外部引用格式
                file_name = os.path.basename(file_path)
                return f"'[{file_name}]{actual_sheet_name}'!{adjusted_ref}"
        return match.group(0)

    def replace_unquoted_match(match):
        sheet_name = match.group(1)
        cell_ref = match.group(2).upper()

        info = find_matching_info(sheet_name)
        if info:
            file_path = info['file_path']
            actual_sheet_name = info['sheet_name']
            adjusted_ref = adjust_cell_ref(cell_ref)

            # 检查是否是本地引用（输出文件自身或模板自引用）
            is_local = (
                (output_file_path and os.path.normpath(file_path) == os.path.normpath(output_file_path)) or
                info.get('is_template_self_reference', False)
            )

            if is_local:
                # 返回本地引用格式
                # 如果 sheet 名需要引号，使用引号
                if any(c in actual_sheet_name for c in " -()&^%$#@!~`'\"\\"):
                    return f"'{actual_sheet_name}'!{adjusted_ref}"
                else:
                    return f"{actual_sheet_name}!{adjusted_ref}"
            else:
                # 返回外部引用格式
                file_name = os.path.basename(file_path)
                return f"'[{file_name}]{actual_sheet_name}'!{adjusted_ref}"
        return match.group(0)

    def replace_bracket_match(match):
        """处理 [xxx]SheetName!CellRef 格式（Excel外部引用索引格式）"""
        full_reference = match.group(1)  # 例如: [3]合同汇总分析表
        cell_ref = match.group(2).upper()

        info = find_matching_info(full_reference)
        if info:
            file_path = info['file_path']
            actual_sheet_name = info['sheet_name']
            adjusted_ref = adjust_cell_ref(cell_ref)

            # 检查是否是本地引用（输出文件自身或模板自引用）
            is_local = (
                (output_file_path and os.path.normpath(file_path) == os.path.normpath(output_file_path)) or
                info.get('is_template_self_reference', False)
            )

            if is_local:
                # 返回本地引用格式
                return f"'{actual_sheet_name}'!{adjusted_ref}"
            else:
                # 返回外部引用格式
                file_name = os.path.basename(file_path)
                return f"'[{file_name}]{actual_sheet_name}'!{adjusted_ref}"
        return match.group(0)

    # 先处理带单引号的sheet名
    result = re.sub(quoted_pattern, replace_quoted_match, formula, flags=re.IGNORECASE)

    # 处理带方括号但无单引号的格式（Excel外部引用索引格式）
    result = re.sub(bracket_pattern, replace_bracket_match, result, flags=re.IGNORECASE)

    # 最后处理不带单引号和方括号的sheet名
    result = re.sub(unquoted_pattern, replace_unquoted_match, result, flags=re.IGNORECASE)

    # 调整本地单元格引用的行号（不包含sheet引用的）
    def adjust_local_ref(match):
        col = match.group(1)
        row = int(match.group(2)) + row_offset
        return f"{col}{row}"

    # 匹配本地单元格引用（不在单引号内，不跟在sheet名后面的）
    local_pattern = r"(?<![A-Za-z!'\"\\])([A-Z]+)(\d+)(?![A-Za-z])"
    result = re.sub(local_pattern, adjust_local_ref, result)

    return result


def apply_template_styles(
    output_ws: openpyxl.worksheet.worksheet.Worksheet,
    template_ws: openpyxl.worksheet.worksheet.Worksheet,
    column_names: List[str],
    max_data_row: int
) -> None:
    """
    将模板的样式应用到输出文件

    Args:
        output_ws: 输出工作表对象
        template_ws: 模板工作表对象
        column_names: 列名列表
        max_data_row: 数据的最大行数
    """
    print("正在应用模板样式...")

    # 应用标题行样式（第一行）
    for col_idx, col_name in enumerate(column_names, start=1):
        template_cell = template_ws.cell(row=1, column=col_idx)
        output_cell = output_ws.cell(row=1, column=col_idx)

        if template_cell.has_style:
            copy_cell_style(template_cell, output_cell)

    # 应用第二行的样式（数据样式模板）
    for col_idx, col_name in enumerate(column_names, start=1):
        template_cell = template_ws.cell(row=2, column=col_idx)

        # 应用到所有数据行
        for row_idx in range(2, max_data_row + 1):
            output_cell = output_ws.cell(row=row_idx, column=col_idx)

            if template_cell.has_style:
                copy_cell_style(template_cell, output_cell)

    # 复制列宽
    for col_idx in range(1, template_ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        if template_ws.column_dimensions[col_letter].width:
            output_ws.column_dimensions[col_letter].width = template_ws.column_dimensions[col_letter].width

    # 复制行高（标题行）
    if template_ws.row_dimensions[1].height:
        output_ws.row_dimensions[1].height = template_ws.row_dimensions[1].height

    print("✓ 模板样式应用完成")


def apply_formulas_to_output(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    formula_columns: List[str],
    formula_templates: Dict[str, str],
    alias_to_info: Dict[str, Dict[str, str]],
    start_row: int = 2,
    output_file_path: Optional[str] = None,
    external_links: Optional[Dict[int, str]] = None
) -> None:
    """
    将公式应用到输出文件

    Args:
        ws: openpyxl工作表对象
        formula_columns: 公式列名列表
        formula_templates: 列名到公式模板的映射
        alias_to_info: 别名到文件信息的映射 {alias: {'file_path': str, 'sheet_name': str}}
        start_row: 开始应用公式的行号（默认为2，跳过标题行）
        output_file_path: 输出文件的完整路径，用于检测本地引用
        external_links: 外部链接映射 {索引号: 文件名}
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
            new_formula = replace_sheet_references(formula_template, alias_to_info, row_offset, output_file_path, external_links)

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

    # 读取外部链接映射
    external_links = read_external_links(template_file)

    template_columns, formula_templates, template_ws = read_template_structure(template_file, template_sheet)

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

    # 首先将输出 sheet 本身作为 0 号
    output_sheet_info = {
        'file_path': output_file,  # 输出文件路径
        'sheet_name': '结果'  # 输出 sheet 名称
    }
    alias_to_info['0'] = output_sheet_info

    # 记录模板 sheet 名称到输出 sheet 信息的映射
    # 用于处理模板自引用的情况（模板中引用的是 template_sheet，应该映射到输出文件的 '结果' sheet）
    template_sheet_info = {
        'file_path': output_file,
        'sheet_name': '结果',
        'is_template_self_reference': True  # 标记这是模板自引用的映射
    }
    alias_to_info[template_sheet.lower()] = template_sheet_info

    # 按顺序为数据源分配编号（从1开始）
    for idx, ds in enumerate(data_sources, start=1):
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

        # 构建info字典
        info = {
            'file_path': config.file_path,
            'sheet_name': config.sheet_name
        }

        # 记录数字索引（1, 2, 3...）到文件信息的映射
        alias_to_info[str(idx)] = info

        # 记录别名到文件信息的映射（如果有别名）
        if config.alias:
            alias_to_info[config.alias.lower()] = info

        # 同时记录实际sheet名到文件信息的映射（支持模板中直接使用实际sheet名）
        alias_to_info[config.sheet_name.lower()] = info

    # 打印数据源编号映射
    print(f"   数据源编号映射:")
    print(f"     [0] -> 输出文件: {os.path.basename(output_file)} (sheet: 结果)")
    for idx, ds in enumerate(data_sources, start=1):
        print(f"     [{idx}] -> {os.path.basename(ds['file_path'])} (sheet: {ds['sheet_name']})")

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
            # 找出哪些公式列包含外部引用
            # 外部引用格式包括: sheet0!, sheet1! 或 [N]SheetName! 或 '[file]Sheet'!
            external_ref_columns = []
            local_formula_columns = []

            # 外部引用的正则模式
            external_ref_pattern = r"(sheet\d+!|\[[^\]]+\][^!'\s]+!|'\[[^\]]+\][^']+'\!)"

            for col in formula_columns:
                if col in template_formulas:
                    formula = template_formulas[col]
                    # 检查是否包含外部引用
                    if re.search(external_ref_pattern, formula, re.IGNORECASE):
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
                    start_row=2,
                    output_file_path=output_file,
                    external_links=external_links
                )
            else:
                # 直接数据模式：只应用本地计算公式
                # 使用与前面相同的外部引用正则模式
                external_ref_pattern = r"(sheet\d+!|\[[^\]]+\][^!'\s]+!|'\[[^\]]+\][^']+'\!)"

                for col in formula_columns:
                    if col in template_formulas:
                        formula = template_formulas[col]
                        # 只处理不包含外部引用的公式
                        if not re.search(external_ref_pattern, formula, re.IGNORECASE):
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
                                    # 传入 alias_to_info 以支持 sheet 引用替换
                                    adjusted_formula = replace_sheet_references(
                                        formula, alias_to_info, row_offset, output_file, external_links
                                    )
                                    ws.cell(row=row_idx, column=col_idx).value = adjusted_formula

                                print(f"   已应用公式: {col} = {formula}")

        # 应用模板样式
        apply_template_styles(ws, template_ws, template_columns, len(output_df) + 1)

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

    # 打印最终文件中的公式汇总
    print("\n" + "=" * 70)
    print("最终文件公式汇总")
    print("=" * 70)

    # 重新读取输出文件，打印所有公式
    try:
        wb_output = openpyxl.load_workbook(output_file, data_only=False)
        ws_output = wb_output.active

        # 获取列名
        header_row = 1
        col_names = {}
        for col_idx, cell in enumerate(ws_output[header_row], start=1):
            if cell.value:
                col_names[col_idx] = str(cell.value)

        # 读取第二行的公式（代表所有行的公式模式）
        formulas_in_output = {}
        if ws_output.max_row >= 2:
            for col_idx in range(1, ws_output.max_column + 1):
                cell = ws_output.cell(row=2, column=col_idx)
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    col_name = col_names.get(col_idx, f"列{col_idx}")
                    formulas_in_output[col_name] = cell.value

        if formulas_in_output:
            print(f"\n公式列数: {len(formulas_in_output)}")
            for col_name, formula in formulas_in_output.items():
                # 截断过长的公式
                if len(formula) > 100:
                    formula_display = formula[:100] + "..."
                else:
                    formula_display = formula
                print(f"  {col_name}: {formula_display}")
        else:
            print("\n无公式列（所有数据均为直接值）")

        wb_output.close()
    except Exception as e:
        print(f"   警告: 无法读取输出文件公式: {e}")

    print("=" * 70)

    # 关闭模板工作簿
    template_ws.parent.close()

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
