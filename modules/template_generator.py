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

from modules._template_core import (
    copy_cell_style,
    _copy_worksheet,
    read_external_links,
    read_template_structure,
    replace_sheet_references,
    parse_formula_references,
)


# ==================== 数据类定义 ====================

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


# ==================== 数据源读取函数 ====================

def read_data_source(config: DataSourceConfig, string_columns: Optional[List[str]] = None) -> pd.DataFrame:
    """
    读取数据源文件

    Args:
        config: 数据源配置
        string_columns: 字符串列列表

    Returns:
        pd.DataFrame: 读取的数据
    """
    print(f"正在读取数据源: {os.path.basename(config.file_path)} 的 {config.sheet_name} 工作表...")

    # 读取所有列为字符串，避免自动类型转换
    df = pd.read_excel(
        config.file_path,
        sheet_name=config.sheet_name,
        dtype=str
    )

    # 应用列映射
    df = _apply_column_mappings(df, config.column_mappings)

    # 处理字符串列
    if string_columns:
        df = _process_string_columns(df, string_columns)

    print(f"数据源 '{config.alias}' 数据: {len(df)} 行, {len(df.columns)} 列")

    return df


def _apply_column_mappings(df: pd.DataFrame, mappings: List[DataColumnMapping]) -> pd.DataFrame:
    """应用列映射"""
    column_map = {m.source_column: m.target_column for m in mappings}
    renamed_columns = {}

    for col in df.columns:
        if col in column_map:
            renamed_columns[col] = column_map[col]

    return df.rename(columns=renamed_columns)


def _process_string_columns(df: pd.DataFrame, string_columns: List[str]) -> pd.DataFrame:
    """处理字符串列"""
    for col in string_columns:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(x) if pd.notna(x) else x)
    return df


# ==================== 数据合并函数 ====================

def merge_data_by_row(
    data_sources: List[Tuple[str, pd.DataFrame]],
    template_columns: List[str]
) -> pd.DataFrame:
    """
    按行号对齐合并多个数据源

    Args:
        data_sources: 数据源列表 [(别名, DataFrame), ...]
        template_columns: 模板列名列表

    Returns:
        pd.DataFrame: 合并后的数据
    """
    if not data_sources:
        return pd.DataFrame(columns=template_columns)

    # 找出最大行数
    max_rows = max(len(df) for _, df in data_sources)
    print(f"合并数据: 最大行数 = {max_rows}")

    # 创建合并后的DataFrame
    merged_data = {col: [None] * max_rows for col in template_columns}

    for alias, df in data_sources:
        for col in template_columns:
            if col in df.columns:
                for i, value in enumerate(df[col]):
                    if i < max_rows:
                        merged_data[col][i] = value

    result = pd.DataFrame(merged_data)
    print(f"合并完成: {len(result)} 行, {len(result.columns)} 列")

    return result


# ==================== 模板样式应用函数 ====================

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

    _apply_header_styles(output_ws, template_ws, column_names)
    _apply_data_styles(output_ws, template_ws, column_names, max_data_row)
    _copy_column_widths(output_ws, template_ws)
    _copy_row_height(output_ws, template_ws)

    print("✓ 模板样式应用完成")


def _apply_header_styles(output_ws, template_ws, column_names: List[str]) -> None:
    """应用标题行样式"""
    for col_idx, col_name in enumerate(column_names, start=1):
        template_cell = template_ws.cell(row=1, column=col_idx)
        output_cell = output_ws.cell(row=1, column=col_idx)

        if template_cell.has_style:
            copy_cell_style(template_cell, output_cell)


def _apply_data_styles(output_ws, template_ws, column_names: List[str], max_data_row: int) -> None:
    """应用数据行样式"""
    for col_idx, col_name in enumerate(column_names, start=1):
        template_cell = template_ws.cell(row=2, column=col_idx)

        for row_idx in range(2, max_data_row + 1):
            output_cell = output_ws.cell(row=row_idx, column=col_idx)

            if template_cell.has_style:
                copy_cell_style(template_cell, output_cell)


def _copy_column_widths(output_ws, template_ws) -> None:
    """复制列宽"""
    for col_idx in range(1, template_ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        if template_ws.column_dimensions[col_letter].width:
            output_ws.column_dimensions[col_letter].width = template_ws.column_dimensions[col_letter].width


def _copy_row_height(output_ws, template_ws) -> None:
    """复制行高（标题行）"""
    if template_ws.row_dimensions[1].height:
        output_ws.row_dimensions[1].height = template_ws.row_dimensions[1].height


# ==================== 公式应用函数 ====================

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
        ws: 输出工作表对象
        formula_columns: 公式列名列表
        formula_templates: 公式模板字典
        alias_to_info: sheet名到文件信息的映射
        start_row: 开始行号
        output_file_path: 输出文件路径
        external_links: 外部链接映射
    """
    # 获取列名到列号的映射
    col_names = _get_column_names(ws)

    for col_name in formula_columns:
        if col_name not in formula_templates:
            continue

        formula = formula_templates[col_name]
        col_idx = col_names.get(col_name)

        if col_idx:
            _apply_formula_to_column(ws, col_idx, formula, alias_to_info,
                                    start_row, output_file_path, external_links, col_name)


def _get_column_names(ws) -> Dict[str, int]:
    """获取列名到列号的映射"""
    col_names = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value:
            col_names[str(cell.value)] = col_idx
    return col_names


def _apply_formula_to_column(ws, col_idx: int, formula: str, alias_to_info: Dict,
                             start_row: int, output_file_path: Optional[str],
                             external_links: Optional[Dict], col_name: str) -> None:
    """应用公式到整列"""
    for row_idx in range(start_row, ws.max_row + 1):
        row_offset = row_idx - start_row
        adjusted_formula = replace_sheet_references(
            formula, alias_to_info, row_offset, output_file_path, external_links
        )
        ws.cell(row=row_idx, column=col_idx).value = adjusted_formula

    print(f"   已应用公式: {col_name} = {formula}")


# ==================== 主函数：生成Excel ====================

def generate_excel_from_template(
    template_file: str,
    template_sheet: str,
    formula_columns: List[str],
    data_sources: List[Dict],
    output_file: str,
    string_columns: Optional[List[str]] = None,
    use_external_refs: bool = False,
    primary_column: Optional[str] = None
) -> pd.DataFrame:
    """
    基于模板生成Excel文件

    Args:
        template_file: 模板文件路径
        template_sheet: 模板sheet名称
        formula_columns: 公式列集合
        data_sources: 数据源集合
        output_file: 输出文件路径
        string_columns: 字符串列列表
        use_external_refs: 是否使用外部引用公式
            - False（默认）: 将数据源sheet添加到输出文件，公式直接引用sheet名
            - True: 公式使用外部文件引用，不复制数据源sheet
        primary_column: 主键列名，为空时跳过过滤

    Returns:
        pd.DataFrame: 生成的数据
    """
    # 确保输出文件路径是绝对路径
    if not os.path.isabs(output_file):
        output_file = os.path.join(os.getcwd(), output_file)

    print("=== Excel模板生成器 ===\n")

    # 1. 验证文件
    _validate_input_files(template_file, data_sources)

    # 2. 分析模板
    external_links, template_formulas, template_columns, template_ws, template_wb = _analyze_template(
        template_file, template_sheet, formula_columns
    )

    # 3. 加载数据源
    loaded_data_sources, alias_to_info = _load_all_data_sources(
        data_sources, output_file, template_sheet, string_columns
    )

    # 4. 合并数据
    merged_df = merge_data_by_row(loaded_data_sources, template_columns)

    # 5. 过滤数据
    merged_df = _filter_data_by_primary_column(merged_df, primary_column)

    # 6. 生成输出文件
    _generate_output_file(
        output_file, merged_df, template_columns, template_formulas,
        formula_columns, alias_to_info, external_links, template_ws,
        use_external_refs, string_columns, data_sources
    )

    # 7. 打印公式汇总
    _print_formula_summary(output_file)

    # 关闭模板工作簿
    template_wb.close()

    return merged_df


def _validate_input_files(template_file: str, data_sources: List[Dict]) -> None:
    """验证输入文件是否存在"""
    print("1. 验证文件...")

    if not os.path.exists(template_file):
        raise FileNotFoundError(f"模板文件不存在: {template_file}")

    for ds in data_sources:
        if not os.path.exists(ds['file_path']):
            raise FileNotFoundError(f"数据文件不存在: {ds['file_path']}")

    print("   所有文件验证通过\n")


def _analyze_template(template_file: str, template_sheet: str,
                      formula_columns: List[str]) -> Tuple:
    """分析模板结构"""
    print("2. 分析模板...")

    external_links = read_external_links(template_file)
    template_columns, formula_templates, template_ws, template_wb = read_template_structure(
        template_file, template_sheet
    )

    # 收集模板中的公式
    template_formulas = {}
    for col_name in formula_columns:
        if col_name in formula_templates:
            template_formulas[col_name] = formula_templates[col_name]
        else:
            print(f"   警告: 公式列 '{col_name}' 在模板中没有公式")

    print()
    return external_links, template_formulas, template_columns, template_ws, template_wb


def _load_all_data_sources(data_sources: List[Dict], output_file: str,
                           template_sheet: str, string_columns: Optional[List[str]]) -> Tuple:
    """加载所有数据源"""
    print("3. 加载数据源...")

    loaded_data_sources = []
    alias_to_info = _init_alias_to_info(output_file, template_sheet)

    for idx, ds in enumerate(data_sources, start=1):
        config = _create_data_source_config(ds)
        df = read_data_source(config, string_columns)
        loaded_data_sources.append((config.alias, df))

        _update_alias_mapping(alias_to_info, str(idx), config)

    _print_data_source_mapping(output_file, data_sources)
    print()

    return loaded_data_sources, alias_to_info


def _init_alias_to_info(output_file: str, template_sheet: str) -> Dict:
    """初始化别名映射"""
    alias_to_info = {}

    # 输出 sheet 作为 0 号
    alias_to_info['0'] = {
        'file_path': output_file,
        'sheet_name': '结果'
    }

    # 模板 sheet 映射到输出 sheet
    alias_to_info[template_sheet.lower()] = {
        'file_path': output_file,
        'sheet_name': '结果',
        'is_template_self_reference': True
    }

    return alias_to_info


def _create_data_source_config(ds: Dict) -> DataSourceConfig:
    """创建数据源配置"""
    column_mappings = [
        DataColumnMapping(source_column=m['source'], target_column=m['target'])
        for m in ds['column_mappings']
    ]

    return DataSourceConfig(
        file_path=ds['file_path'],
        sheet_name=ds['sheet_name'],
        column_mappings=column_mappings,
        alias=ds.get('alias', '')
    )


def _update_alias_mapping(alias_to_info: Dict, idx: str, config: DataSourceConfig) -> None:
    """更新别名映射"""
    info = {
        'file_path': config.file_path,
        'sheet_name': config.sheet_name
    }

    alias_to_info[idx] = info

    if config.alias:
        alias_to_info[config.alias.lower()] = info

    alias_to_info[config.sheet_name.lower()] = info


def _print_data_source_mapping(output_file: str, data_sources: List[Dict]) -> None:
    """打印数据源映射"""
    print(f"   数据源编号映射:")
    print(f"     [0] -> 输出文件: {os.path.basename(output_file)} (sheet: 结果)")

    for idx, ds in enumerate(data_sources, start=1):
        print(f"     [{idx}] -> {os.path.basename(ds['file_path'])} (sheet: {ds['sheet_name']})")


def _filter_data_by_primary_column(merged_df: pd.DataFrame,
                                    primary_column: Optional[str]) -> pd.DataFrame:
    """根据主键列过滤数据"""
    if not primary_column:
        return merged_df

    if primary_column not in merged_df.columns:
        print(f"   警告: 主键列 '{primary_column}' 不存在于数据中，跳过过滤")
        return merged_df

    original_count = len(merged_df)
    merged_df = merged_df[
        merged_df[primary_column].notna() &
        (merged_df[primary_column].astype(str).str.strip() != '')
    ]

    filtered_count = len(merged_df)
    if original_count > filtered_count:
        print(f"   过滤掉 {original_count - filtered_count} 行（{primary_column} 列为空）")

    return merged_df


def _generate_output_file(output_file: str, merged_df: pd.DataFrame,
                          template_columns: List[str], template_formulas: Dict,
                          formula_columns: List[str], alias_to_info: Dict,
                          external_links: Dict, template_ws,
                          use_external_refs: bool, string_columns: Optional[List[str]],
                          data_sources: Optional[List[Dict]] = None) -> None:
    """生成输出文件"""
    print("5. 生成输出文件...")

    if use_external_refs:
        print("   模式: 外部引用公式（数据源保留在外部文件）")
    else:
        print("   模式: 内部引用公式（数据源sheet将复制到输出文件）")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        output_df = merged_df.copy()

        # 准备输出DataFrame
        local_formula_columns = _prepare_output_df(
            output_df, template_formulas, formula_columns, use_external_refs
        )

        output_df.to_excel(writer, sheet_name='结果', index=False)
        ws = writer.sheets['结果']

        # 如果不使用外部引用，复制数据源sheet到输出文件
        if not use_external_refs and data_sources:
            _copy_data_source_sheets(writer, data_sources, alias_to_info)

        # 应用公式
        _apply_formulas(
            ws, output_df, template_formulas, formula_columns,
            alias_to_info, output_file, external_links, use_external_refs, local_formula_columns
        )

        # 应用模板样式
        apply_template_styles(ws, template_ws, template_columns, len(output_df) + 1)

        # 处理字符串列格式
        if string_columns:
            _apply_string_column_format(ws, output_df, string_columns)

    print(f"\n输出文件已保存: {output_file}")


def _copy_data_source_sheets(writer, data_sources: List[Dict], alias_to_info: Dict) -> None:
    """
    将数据源sheet复制到输出文件

    Args:
        writer: ExcelWriter对象
        data_sources: 数据源列表
        alias_to_info: 别名到信息的映射
    """
    print("   正在复制数据源sheet到输出文件...")

    copied_sheets = set()  # 跟踪已复制的sheet，避免重复

    for idx, ds in enumerate(data_sources, start=1):
        file_path = ds['file_path']
        sheet_name = ds['sheet_name']
        alias = ds.get('alias', '')

        # 确定目标sheet名称
        if alias:
            target_sheet_name = alias
        else:
            target_sheet_name = sheet_name

        # 避免sheet名冲突
        if target_sheet_name in copied_sheets or target_sheet_name == '结果':
            target_sheet_name = f"{target_sheet_name}_{idx}"

        # 读取数据源文件
        try:
            source_wb = openpyxl.load_workbook(file_path, data_only=False)
            if sheet_name not in source_wb.sheetnames:
                print(f"   警告: 数据源 {file_path} 中不存在sheet '{sheet_name}'")
                source_wb.close()
                continue

            source_ws = source_wb[sheet_name]

            # 创建新sheet
            target_ws = writer.book.create_sheet(title=target_sheet_name)

            # 复制所有单元格数据和样式
            _copy_worksheet(source_ws, target_ws)

            # 更新alias_to_info，使其指向内部的sheet
            info_update = {
                'file_path': '',  # 空路径表示内部引用
                'sheet_name': target_sheet_name,
                'is_internal': True
            }

            # 更新所有相关映射
            alias_to_info[str(idx)]['is_internal'] = True
            alias_to_info[str(idx)]['sheet_name'] = target_sheet_name
            alias_to_info[str(idx)]['file_path'] = ''

            if alias:
                if alias.lower() in alias_to_info:
                    alias_to_info[alias.lower()]['is_internal'] = True
                    alias_to_info[alias.lower()]['sheet_name'] = target_sheet_name
                    alias_to_info[alias.lower()]['file_path'] = ''

            if sheet_name.lower() in alias_to_info:
                alias_to_info[sheet_name.lower()]['is_internal'] = True
                alias_to_info[sheet_name.lower()]['sheet_name'] = target_sheet_name
                alias_to_info[sheet_name.lower()]['file_path'] = ''

            copied_sheets.add(target_sheet_name)
            print(f"   已复制: {os.path.basename(file_path)}[{sheet_name}] -> [{target_sheet_name}]")

            source_wb.close()
        except Exception as e:
            print(f"   警告: 无法复制数据源sheet '{sheet_name}': {e}")


def _prepare_output_df(output_df: pd.DataFrame, template_formulas: Dict,
                        formula_columns: List[str], use_external_refs: bool) -> List[str]:
    """
    准备输出DataFrame

    Args:
        output_df: 输出DataFrame
        template_formulas: 模板公式字典
        formula_columns: 公式列列表
        use_external_refs: 是否使用外部引用

    Returns:
        List[str]: 需要应用公式的列名列表
    """
    # 收集所有需要应用公式的列
    applicable_formula_columns = []

    for col in formula_columns:
        if col in template_formulas:
            applicable_formula_columns.append(col)
            # 清空公式列，为后续写入公式做准备
            if col in output_df.columns:
                output_df[col] = None

    print(f"   公式列: {applicable_formula_columns}")

    return applicable_formula_columns


def _apply_formulas(ws, output_df: pd.DataFrame, template_formulas: Dict,
                    formula_columns: List[str], alias_to_info: Dict,
                    output_file: str, external_links: Dict,
                    use_external_refs: bool, formula_columns_to_apply: List[str]) -> None:
    """
    应用公式到工作表

    公式引用格式由 alias_to_info 中的 is_internal 标志决定：
    - is_internal=True: 内部引用（SheetName!A1）
    - is_internal=False: 外部引用（[filename]SheetName!A1）
    """
    if not formula_columns_to_apply or not template_formulas:
        return

    apply_formulas_to_output(
        ws, formula_columns_to_apply, template_formulas, alias_to_info,
        start_row=2, output_file_path=output_file, external_links=external_links
    )


def _find_column_index(df: pd.DataFrame, col_name: str) -> Optional[int]:
    """查找列索引"""
    for c_idx, c_name in enumerate(df.columns, start=1):
        if c_name == col_name:
            return c_idx
    return None


def _apply_string_column_format(ws, output_df: pd.DataFrame, string_columns: List[str]) -> None:
    """应用字符串列格式"""
    from openpyxl.styles import Font

    for col_idx, col_name in enumerate(output_df.columns, start=1):
        if col_name not in string_columns:
            continue

        ws.column_dimensions[get_column_letter(col_idx)].width = 15

        for row in range(2, len(output_df) + 2):
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = '@'

            if pd.notna(output_df.iloc[row - 2][col_name]):
                cell.value = str(output_df.iloc[row - 2][col_name])


def _print_formula_summary(output_file: str) -> None:
    """打印公式汇总"""
    print("\n" + "=" * 70)
    print("最终文件公式汇总")
    print("=" * 70)

    try:
        wb_output = openpyxl.load_workbook(output_file, data_only=False)
        ws_output = wb_output.active

        col_names = _read_output_column_names(ws_output)
        formulas = _read_output_formulas(ws_output, col_names)

        if formulas:
            print(f"\n公式列数: {len(formulas)}")
            for col_name, formula in formulas.items():
                formula_display = formula[:100] + "..." if len(formula) > 100 else formula
                print(f"  {col_name}: {formula_display}")
        else:
            print("\n无公式列（所有数据均为直接值）")

        wb_output.close()
    except Exception as e:
        print(f"   警告: 无法读取输出文件公式: {e}")

    print("=" * 70)


def _read_output_column_names(ws) -> Dict[int, str]:
    """读取输出文件的列名"""
    col_names = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value:
            col_names[col_idx] = str(cell.value)
    return col_names


def _read_output_formulas(ws, col_names: Dict[int, str]) -> Dict[str, str]:
    """读取输出文件的公式"""
    formulas = {}

    if ws.max_row < 2:
        return formulas

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col_idx)
        if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
            col_name = col_names.get(col_idx, f"列{col_idx}")
            formulas[col_name] = cell.value

    return formulas


# ==================== 命令行解析辅助函数 ====================

def parse_column_mappings(mappings_str: str) -> List[Dict[str, str]]:
    """
    解析列映射字符串

    Args:
        mappings_str: 列映射字符串，格式为 "SourceCol:TargetCol,SourceCol2:TargetCol2"
                      或 "Col1,Col2"（源列名和目标列名相同）

    Returns:
        List[Dict[str, str]]: 列映射列表
    """
    mappings = []

    if not mappings_str:
        return mappings

    pairs = mappings_str.split(',')

    for pair in pairs:
        if ':' in pair:
            source, target = pair.split(':', 1)
            mappings.append({
                'source': source.strip(),
                'target': target.strip()
            })
        else:
            # 不带冒号时，源列名和目标列名相同
            col_name = pair.strip()
            if col_name:
                mappings.append({
                    'source': col_name,
                    'target': col_name
                })

    return mappings
