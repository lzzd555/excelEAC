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
