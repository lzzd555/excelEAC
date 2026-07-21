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


_REF_SHEET_PATTERN = re.compile(r"(?:'([^']+)'!|(?:\[[^\]]+\])?([A-Za-z_][A-Za-z0-9_]*)!)")


def _collect_referenced_sheets(formulas) -> List[str]:
    """从一批公式里收集所有被引用的 sheet 名(去重)。"""
    names = []
    seen = set()
    for f in formulas.values():
        for m in _REF_SHEET_PATTERN.finditer(f or ''):
            name = m.group(1) or m.group(2)
            if name and name.lower() not in seen:
                seen.add(name.lower())
                names.append(name)
    return names


def _sheet_row_count(file_path: str, sheet: str) -> int:
    wb = openpyxl.load_workbook(file_path, read_only=True)
    try:
        ws = wb[sheet]
        return ws.max_row
    finally:
        wb.close()


def resolve_row_source(
    row_source: Optional[Tuple[str, str]],
    formulas: Dict[str, str],
    file_sheet_index: Dict[str, List[str]],
    data_files: Optional[List[str]] = None,
) -> Tuple[int, str]:
    """
    决定输出公式列的行数 N,返回 (N, 说明)。
    1) row_source 显式 → 用 (文件, sheet) 的行数
    2) 否则 → 公式中被引用次数最多的数据 sheet
    3) 再否则 → 第一个数据文件的第一个 sheet
    """
    if data_files is None:
        data_files = list(file_sheet_index.keys())

    # ① 显式
    if row_source:
        f, s = row_source
        f = _match_data_file(f, data_files)
        if s not in file_sheet_index.get(f, []):
            raise ValueError(f"row_source 指定的 sheet '{s}' 不在文件 '{f}' 中")
        return _sheet_row_count(f, s), f"row_source={f}:{s}"

    # ② 统计每个数据 sheet 被引用次数
    refs = _collect_referenced_sheets(formulas)
    counts = {}
    for f in formulas.values():
        for m in _REF_SHEET_PATTERN.finditer(f or ''):
            name = (m.group(1) or m.group(2) or '').lower()
            if name:
                counts[name] = counts.get(name, 0) + 1

    best = None
    best_count = 0
    for name in refs:
        nl = name.lower()
        for f, sheets in file_sheet_index.items():
            for s in sheets:
                if s.lower() == nl:
                    c = counts.get(nl, 0)
                    if c > best_count:
                        best_count = c
                        best = (f, s)
    if best:
        return _sheet_row_count(*best), f"最常引用={best[0]}:{best[1]}"

    # ③ 兜底:第一个数据文件的第一个 sheet
    f = data_files[0]
    s = file_sheet_index[f][0]
    return _sheet_row_count(f, s), f"兜底={f}:{s}"


def _validate_files(template_file: str, data_files: List[str]) -> None:
    if not os.path.exists(template_file):
        raise FileNotFoundError(f"模板文件不存在: {template_file}")
    for f in data_files:
        if not os.path.exists(f):
            raise FileNotFoundError(f"数据文件不存在: {f}")


def _copy_referenced_external_sheets(writer, referenced_infos: List[Dict]) -> None:
    """内部模式:把被引用的外部/模板 sheet 复制进输出文件(去重)。"""
    copied = set()
    for info in referenced_infos:
        if info.get('is_template_self_reference'):
            continue  # 模板自引用指向输出自身,无需复制
        src_file = info['file_path']
        src_sheet = info['sheet_name']
        key = (src_file, src_sheet)
        if not src_file or key in copied or src_sheet in writer.book.sheetnames:
            continue
        src_wb = openpyxl.load_workbook(src_file, data_only=False)
        try:
            if src_sheet not in src_wb.sheetnames:
                continue
            target = writer.book.create_sheet(title=src_sheet)
            _copy_worksheet(src_wb[src_sheet], target)
            copied.add(key)
        finally:
            src_wb.close()


def _apply_formula_column_styles(ws, template_ws, formula_columns: List[str],
                                 col_to_idx: Dict[str, int], n_rows: int) -> None:
    """把模板公式列的表头(第1行)+ 数据行(第2行)样式贴到输出对应列。"""
    tpl_names = [c.value for c in template_ws[1]]
    for col in formula_columns:
        if col not in col_to_idx:
            continue
        out_idx = col_to_idx[col]
        tpl_idx = (tpl_names.index(col) + 1) if col in tpl_names else out_idx
        # 表头样式
        if template_ws.cell(row=1, column=tpl_idx).has_style:
            copy_cell_style(template_ws.cell(row=1, column=tpl_idx), ws.cell(row=1, column=out_idx))
        # 数据行样式(以模板第2行为模板)
        if template_ws.cell(row=2, column=tpl_idx).has_style:
            for r in range(2, n_rows + 2):
                copy_cell_style(template_ws.cell(row=2, column=tpl_idx), ws.cell(row=r, column=out_idx))


def generate_formulas_from_template(
    template_file: str,
    template_sheet: str,
    data_files: List[str],
    output_file: str,
    sheet_mapping: Optional[Dict[str, str]] = None,
    row_source: Optional[Tuple[str, str]] = None,
    use_external_refs: bool = False,
) -> pd.DataFrame:
    """
    仅凭模板 + 数据文件,生成只含公式列的输出。
    返回只含公式列的数据骨架 DataFrame(公式写入 Excel,DataFrame 单元格为空)。
    """
    if not os.path.isabs(output_file):
        output_file = os.path.join(os.getcwd(), output_file)
    print("=== Excel 公式模板生成器 ===\n")

    # 1. 校验
    _validate_files(template_file, data_files)

    # 2. 分析模板(核心原语)
    template_columns, formula_templates, template_ws, template_wb = read_template_structure(
        template_file, template_sheet
    )
    formula_columns = [c for c in template_columns if formula_templates.get(c)]
    if not formula_columns:
        raise ValueError(f"模板 sheet '{template_sheet}' 第 2 行未检测到公式列")

    external_links = read_external_links(template_file)

    # 3. 发现数据文件 sheet
    file_sheet_index = {f: discover_file_sheets(f) for f in data_files}
    template_sheets = template_wb.sheetnames

    # 4. 三层解析公式引用的每个 sheet
    referenced = _collect_referenced_sheets(formula_templates)
    alias_to_info: Dict[str, Dict] = {
        # 模板自引用(template_sheet 引用自己)→ 指向输出 '结果'
        template_sheet.lower(): {'file_path': output_file, 'sheet_name': '结果', 'is_template_self_reference': True},
    }
    referenced_infos = []
    for s in referenced:
        info = resolve_formula_sheet(s, sheet_mapping, file_sheet_index, template_sheets, use_external_refs)
        alias_to_info[s.lower()] = info
        referenced_infos.append(info)

    # 5. 行数
    n_rows, chosen = resolve_row_source(row_source, formula_templates, file_sheet_index, data_files)
    print(f"行数驱动: {chosen} → {n_rows} 行")

    # 6. 骨架 DataFrame(只含公式列)
    output_df = pd.DataFrame({c: [None] * n_rows for c in formula_columns})

    # 7. 写出
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name='结果', index=False)
        ws = writer.sheets['结果']

        # 内部模式:复制被引用的外部 sheet
        if not use_external_refs:
            _copy_referenced_external_sheets(writer, referenced_infos)

        # 逐行写公式(行偏移 = 当前行 - 2)
        col_to_idx = {name: i + 1 for i, name in enumerate(output_df.columns)}
        for col in formula_columns:
            tmpl = formula_templates[col]
            for r in range(2, n_rows + 2):
                formula = replace_sheet_references(
                    tmpl, alias_to_info, row_offset=r - 2,
                    output_file_path=output_file, external_links=external_links,
                )
                ws.cell(row=r, column=col_to_idx[col], value=formula)

        # 8. 公式列样式
        _apply_formula_column_styles(ws, template_ws, formula_columns, col_to_idx, n_rows)

    template_wb.close()
    print(f"\n输出文件已保存: {output_file}")
    return output_df
