"""
Excel 模板公式专用生成模块
仅凭模板 + 数据文件路径,生成「只含公式列」的输出。
与 template_generator(老方法)物理分离,只复用 _template_core 的纯原语。
"""

import os
import re
from typing import List, Dict, Optional, Tuple

import pandas as pd
import openpyxl

from modules._template_core import (
    read_template_structure,
    read_external_links,
    replace_sheet_references,
    copy_cell_style,
    _copy_worksheet,
    _extract_sheet_name,
)


def discover_file_sheets(file_path: str) -> List[str]:
    """列出数据文件中的所有 sheet 名(只读,用完即关)。"""
    wb = openpyxl.load_workbook(file_path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _match_data_file(hint: str, data_files: List[str]) -> str:
    """按完整路径或文件名匹配 data_files 中的某个文件,歧义/未命中则抛错。"""
    hint_norm = os.path.normpath(hint)
    exact = [f for f in data_files if os.path.normpath(f) == hint_norm]
    if len(exact) == 1:
        return exact[0]
    base = os.path.basename(hint_norm)
    by_base = [f for f in data_files if os.path.basename(f) == base]
    if len(by_base) == 1:
        return by_base[0]
    raise ValueError(f"sheet_mapping 中的文件 '{hint}' 无法在 data_files 中唯一匹配: {data_files}")


def _make_sheet_info(sheet_name: str, file_path: str, use_external_refs: bool) -> Dict:
    return {
        'file_path': file_path,
        'sheet_name': sheet_name,
        'is_internal': not use_external_refs,
    }


def resolve_formula_sheet(
    sheet_name: str,
    sheet_mapping: Optional[Dict[str, str]],
    file_sheet_index: Dict[str, List[str]],
    template_sheets: List[str],
    use_external_refs: bool,
) -> Dict:
    """
    三层优先级解析公式中的 sheet 引用:
      ① sheet_mapping 显式映射(最高)
      ② 外部数据文件自动同名匹配
      ③ 模板自带 sheet
    三层均未命中,或第②层同名冲突未消歧 → 抛 ValueError。
    """
    # 去掉 [file] 前缀等,取实际 sheet 名
    actual = _extract_sheet_name(sheet_name)
    key = actual.lower()

    # ① 显式映射
    if sheet_mapping:
        for mkey, mfile in sheet_mapping.items():
            if mkey.lower() == key:
                resolved = _match_data_file(mfile, list(file_sheet_index.keys()))
                if actual not in file_sheet_index.get(resolved, []):
                    raise ValueError(f"sheet_mapping 指定 '{mkey}' → '{mfile}',但该文件中无 sheet '{actual}'")
                return _make_sheet_info(actual, resolved, use_external_refs)

    # ② 外部数据文件自动匹配
    matches = [(f, s) for f, sheets in file_sheet_index.items() for s in sheets if s.lower() == key]
    if len(matches) == 1:
        f, s = matches[0]
        return _make_sheet_info(s, f, use_external_refs)
    if len(matches) > 1:
        files = [m[0] for m in matches]
        raise ValueError(f"sheet '{actual}' 在多个数据文件中均存在: {files};请用 sheet_mapping 指明")

    # ③ 模板自带 sheet
    for s in template_sheets:
        if s.lower() == key:
            return {
                'file_path': '',
                'sheet_name': s,
                'is_template_self_reference': True,
            }

    raise ValueError(f"公式引用 sheet '{actual}':映射/数据文件/模板中均未找到;可用 sheet_mapping 指定")
