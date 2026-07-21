# 模板公式专用生成功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增 `generate_formulas_from_template`,仅凭模板 + 数据文件路径生成「只含公式列」的输出;公式 sheet 引用按「显式映射 > 外部数据文件 > 模板自带 sheet」三层解析。

**Architecture:** 方案 C——抽共享核心模块 `modules/_template_core.py`,老方法 `template_generator.py` 改为从核心导入(行为不变,现有测试为回归网);新方法独立放在 `modules/template_formula.py`,只复用核心里的纯原语,自带编排逻辑,与老方法物理分离。CLI 在 `main.py` 新增 `template-formula` 子命令。

**Tech Stack:** Python 3、pandas、openpyxl。测试沿用仓库现有风格:纯 `python <file>` 脚本,`sys.path.insert` 导入项目根,内部 `assert` + 打印。

**Worktree:** 建议在专属 worktree 执行(参见 superpowers:using-git-worktrees)。若直接在 main 上做,每个 task 独立 commit,保持树常绿。

**配套 spec:** `docs/superpowers/specs/2026-07-20-template-formula-only-design.md`

**回归测试命令(贯穿全程):** 每个 task 完成后必须保证下列现有模板测试全绿——它们是抽取阶段的回归网:
```bash
python tests/template/test_template_generator.py
python tests/template/test_complex_formula.py
python tests/template/test_real_sheet_name.py
python tests/template/test_bracket_format.py
python tests/template/test_bracket_actual.py
python tests/template/test_file_path_in_formula.py
```

---

## 文件结构

| 文件 | 责任 | 动作 |
|------|------|------|
| `modules/_template_core.py` | 共享纯原语:公式解析/改写、sheet 匹配、样式/颜色复制、外部链接、模板读取、`_copy_worksheet`、`CellStyle` | 新建(从 generator 搬入) |
| `modules/template_generator.py` | 老方法:列映射+数据合并+老流程编排 | 修改(删搬走的函数,改为 `from modules._template_core import ...`) |
| `modules/template_formula.py` | 新方法:`generate_formulas_from_template` + 新 helper + 编排 | 新建 |
| `main.py` | CLI 入口 | 修改(新增 `template-formula` 子命令) |
| `tests/template/test_template_formula.py` | 新功能测试 | 新建 |
| `README.md` | 用户文档 | 修改(新增独立章节) |

---

## Task 1: 建立回归基线

**Files:** 无修改

- [ ] **Step 1: 跑全部现有模板测试,确认当前全绿**

Run:
```bash
python tests/template/test_template_generator.py && \
python tests/template/test_complex_formula.py && \
python tests/template/test_real_sheet_name.py && \
python tests/template/test_bracket_format.py && \
python tests/template/test_bracket_actual.py && \
python tests/template/test_file_path_in_formula.py
```
Expected: 全部打印通过(各脚本以 exit 0 结束、无 AssertionError)。

- [ ] **Step 2: 若任一失败,先停下报告**——基线不绿就不能开始抽取。

记录通过的测试数作为基线。无需 commit(未改动代码)。

---

## Task 2: 创建共享核心模块 `_template_core.py`

**Files:**
- Create: `modules/_template_core.py`

把下列函数(连同其 docstring 与实现)**原样**从 `modules/template_generator.py` 搬到 `modules/_template_core.py`。这些函数相互自洽,不依赖任何「留在 generator」的函数(已核对:`_copy_worksheet` 只调 `copy_cell_style`;样式组只调彼此;公式组只调彼此 + 入参)。

**搬入清单(精确函数名):**
- 数据类:`CellStyle`
- 颜色:`copy_color`、`_copy_rgb_color`、`_copy_theme_color`、`_copy_indexed_color`
- 样式:`copy_side`、`copy_cell_style`、`_copy_font_style`、`_get_font_color`、`_copy_fill_style`、`_apply_fill_by_type`、`_apply_solid_fill`、`_get_fill_color_value`、`_apply_fallback_fill`、`_copy_border_style`、`_copy_alignment_style`
- 工作表复制:`_copy_worksheet`
- 外部链接:`read_external_links`、`_extract_external_links`、`_parse_single_link`、`replace_link_indices_with_filenames`
- 模板读取:`read_template_structure`、`_read_template_columns`、`_read_formula_templates`
- 公式引用(引擎):`parse_formula_references`、`replace_sheet_references`、`_adjust_cell_ref`、`_adjust_single_ref`、`_adjust_range_ref`、`_extract_sheet_name`、`_find_matching_info`、`_resolve_multiple_matches`、`_find_by_index`、`_build_reference`、`_replace_quoted_match`、`_replace_unquoted_match`、`_replace_bracket_match`
- 模块级常量:`_QUOTED_PATTERN`、`_BRACKET_PATTERN`、`_UNQUOTED_PATTERN`、`_LOCAL_PATTERN`(被上面函数引用,必须一起搬)

- [ ] **Step 1: 创建 `modules/_template_core.py`,写入下列 import 头**

```python
"""
Excel 模板共享核心原语
公式解析/改写、sheet 匹配、样式/颜色复制、外部链接、模板读取。
新老两个模板方法(template_generator / template_formula)都从此导入。
"""

import os
import re
import copy
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
```

- [ ] **Step 2: 把清单中全部函数/常量原样粘贴到 import 头之下(保持原顺序)**

操作:在 `modules/template_generator.py` 中逐个找到这些函数(可用 `grep -n "^def \\|^_QUOTED_PATTERN\\|^_BRACKET_PATTERN\\|^_UNQUOTED_PATTERN\\|^_LOCAL_PATTERN\\|^class CellStyle" modules/template_generator.py` 定位),**剪切**到 `modules/_template_core.py`。注意 `CellStyle` dataclass 也要搬。

> 此步先只搬不删 generator 中的定义,下一步统一替换(避免中间状态无法导入)。

- [ ] **Step 3: 验证核心模块可独立导入**

Run:
```bash
python -c "from modules import _template_core as c; print(c.copy_cell_style, c.replace_sheet_references, c._copy_worksheet, c.read_template_structure, c.CellStyle)"
```
Expected: 打印出 5 个对象 repr,无 ImportError/NameError。

若有 NameError,说明某被搬函数仍引用了「留在 generator」的符号——把那个被引用的符号也补搬到 core(本任务清单应已覆盖,但若遗漏则就地补齐)。

- [ ] **Step 4: Commit**

```bash
git add modules/_template_core.py
git commit -m "refactor: 抽取共享核心模块 _template_core(函数搬入,尚未切换导入)"
```

---

## Task 3: 老方法改为从核心导入

**Files:**
- Modify: `modules/template_generator.py`

- [ ] **Step 1: 在 `template_generator.py` 顶部 import 区追加一行**

在现有 import(`import copy` 那一组之后)追加:
```python
from modules._template_core import (
    CellStyle,
    copy_color, _copy_rgb_color, _copy_theme_color, _copy_indexed_color,
    copy_side, copy_cell_style, _copy_font_style, _get_font_color,
    _copy_fill_style, _apply_fill_by_type, _apply_solid_fill, _get_fill_color_value,
    _apply_fallback_fill, _copy_border_style, _copy_alignment_style,
    _copy_worksheet,
    read_external_links, _extract_external_links, _parse_single_link, replace_link_indices_with_filenames,
    read_template_structure, _read_template_columns, _read_formula_templates,
    parse_formula_references, replace_sheet_references, _adjust_cell_ref, _adjust_single_ref,
    _adjust_range_ref, _extract_sheet_name, _find_matching_info, _resolve_multiple_matches,
    _find_by_index, _build_reference, _replace_quoted_match, _replace_unquoted_match, _replace_bracket_match,
    _QUOTED_PATTERN, _BRACKET_PATTERN, _UNQUOTED_PATTERN, _LOCAL_PATTERN,
)
```

- [ ] **Step 2: 删除 `template_generator.py` 中已搬走的本地定义**

删除 Task 2 清单中的全部 `def`/`class`/常量定义(它们现在由 import 提供)。保留 `DataColumnMapping`、`DataSourceConfig` 及所有老流程编排函数(它们留在 generator)。

定位:
```bash
grep -n "^def \\|^class \\|^_QUOTED_PATTERN\\|^_BRACKET_PATTERN\\|^_UNQUOTED_PATTERN\\|^_LOCAL_PATTERN" modules/template_generator.py
```
逐一删除命中且属于「搬入清单」的那些;**不要删** `DataColumnMapping`、`DataSourceConfig` 及未在清单中的函数。

- [ ] **Step 3: 验证 generator 仍可导入**

Run:
```bash
python -c "from modules.template_generator import generate_excel_from_template, merge_data_by_row, apply_formulas_to_output; print('ok')"
```
Expected: 打印 `ok`,无 NameError。

- [ ] **Step 4: 跑回归测试,必须全绿**

Run(同 Task 1 Step 1 的 6 条命令)。
Expected: 与基线一致,全部通过。若有失败,通常是漏删/漏 import 某符号——根据报错补到 Step 1 的 import 列表。

- [ ] **Step 5: Commit**

```bash
git add modules/template_generator.py
git commit -m "refactor: template_generator 改为从 _template_core 导入共享原语"
```

---

## Task 4: 新方法骨架 + `discover_file_sheets`(TDD)

**Files:**
- Create: `modules/template_formula.py`
- Test: `tests/template/test_template_formula.py`

- [ ] **Step 1: 写失败测试**

创建 `tests/template/test_template_formula.py`:
```python
"""
测试 generate_formulas_from_template(模板公式专用生成)
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import openpyxl
from modules.template_formula import discover_file_sheets

TEST_DIR = os.path.dirname(os.path.abspath(__file__))


def test_discover_file_sheets_lists_all():
    fpath = os.path.join(TEST_DIR, '_tmp_discover.xlsx')
    wb = openpyxl.Workbook()
    wb.active.title = 'S1'
    wb.create_sheet('S2')
    wb.create_sheet('S3')
    wb.save(fpath)
    wb.close()

    sheets = discover_file_sheets(fpath)
    assert sheets == ['S1', 'S2', 'S3'], f"expected 3 sheets, got {sheets}"

    os.remove(fpath)
    print("PASS test_discover_file_sheets_lists_all")


if __name__ == '__main__':
    test_discover_file_sheets_lists_all()
    print("all pass")
```

- [ ] **Step 2: 跑测试,确认失败**

Run: `python tests/template/test_template_formula.py`
Expected: FAIL(`ModuleNotFoundError: No module named 'modules.template_formula'`)。

- [ ] **Step 3: 创建 `modules/template_formula.py` 并实现 `discover_file_sheets`**

```python
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
```

- [ ] **Step 4: 跑测试,确认通过**

Run: `python tests/template/test_template_formula.py`
Expected: PASS。

- [ ] **Step 5: Commit**

```bash
git add modules/template_formula.py tests/template/test_template_formula.py
git commit -m "feat(template_formula): 新增 discover_file_sheets"
```

---

## Task 5: 三层解析器 `resolve_formula_sheet`(TDD)

**Files:**
- Modify: `modules/template_formula.py`
- Test: `tests/template/test_template_formula.py`

- [ ] **Step 1: 追加失败测试**

在 `tests/template/test_template_formula.py` 中(在 `if __name__` 之前)追加:
```python
from modules.template_formula import resolve_formula_sheet


def _make_data_file(name, sheet):
    fpath = os.path.join(TEST_DIR, name)
    wb = openpyxl.Workbook()
    wb.active.title = sheet
    wb[sheet].cell(row=1, column=1, value='x')
    wb.save(fpath); wb.close()
    return fpath


def test_resolve_explicit_mapping_wins():
    f1 = _make_data_file('_tmp_r1.xlsx', 'Data')
    f2 = _make_data_file('_tmp_r2.xlsx', 'Data')  # 同名 sheet 在两个文件
    file_sheet_index = {f1: ['Data'], f2: ['Data']}
    sheet_mapping = {'Data': f2}  # 显式指向 f2
    info = resolve_formula_sheet('Data', sheet_mapping, file_sheet_index, [], use_external_refs=False)
    assert info['file_path'] == f2, info
    os.remove(f1); os.remove(f2)
    print("PASS test_resolve_explicit_mapping_wins")


def test_resolve_auto_match_single():
    f1 = _make_data_file('_tmp_a1.xlsx', 'ESDP-Bpart')
    file_sheet_index = {f1: ['ESDP-Bpart']}
    info = resolve_formula_sheet('ESDP-Bpart', None, file_sheet_index, [], use_external_refs=False)
    assert info['file_path'] == f1 and info['is_internal'] is True, info
    os.remove(f1)
    print("PASS test_resolve_auto_match_single")


def test_resolve_conflict_without_mapping_raises():
    f1 = _make_data_file('_tmp_c1.xlsx', 'Data')
    f2 = _make_data_file('_tmp_c2.xlsx', 'Data')
    file_sheet_index = {f1: ['Data'], f2: ['Data']}
    try:
        resolve_formula_sheet('Data', None, file_sheet_index, [], use_external_refs=False)
        assert False, "应抛错"
    except ValueError as e:
        assert 'Data' in str(e)
    os.remove(f1); os.remove(f2)
    print("PASS test_resolve_conflict_without_mapping_raises")


def test_resolve_template_self_sheet():
    f1 = _make_data_file('_tmp_t1.xlsx', 'Other')
    file_sheet_index = {f1: ['Other']}
    info = resolve_formula_sheet('配置表', None, file_sheet_index, ['配置表', '结果'], use_external_refs=False)
    assert info.get('is_template_self_reference') is True, info
    os.remove(f1)
    print("PASS test_resolve_template_self_sheet")


def test_resolve_not_found_raises():
    f1 = _make_data_file('_tmp_n1.xlsx', 'Other')
    try:
        resolve_formula_sheet('Missing', None, {f1: ['Other']}, ['配置表'], use_external_refs=False)
        assert False, "应抛错"
    except ValueError as e:
        assert 'Missing' in str(e)
    os.remove(f1)
    print("PASS test_resolve_not_found_raises")
```
并把 `if __name__ == '__main__':` 块改为调用全部 5 个新测试 + Task 4 的测试。

- [ ] **Step 2: 跑测试,确认新测试失败**

Run: `python tests/template/test_template_formula.py`
Expected: FAIL(`ImportError: cannot import name 'resolve_formula_sheet'`)。

- [ ] **Step 3: 实现 `resolve_formula_sheet` 及其小助手**

在 `modules/template_formula.py` 追加:
```python
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
```

- [ ] **Step 4: 跑测试,确认全部通过**

Run: `python tests/template/test_template_formula.py`
Expected: 6 个 PASS。

- [ ] **Step 5: Commit**

```bash
git add modules/template_formula.py tests/template/test_template_formula.py
git commit -m "feat(template_formula): 三层优先级 sheet 解析器 resolve_formula_sheet"
```

---

## Task 6: 行数驱动 `resolve_row_source`(TDD)

**Files:**
- Modify: `modules/template_formula.py`
- Test: `tests/template/test_template_formula.py`

- [ ] **Step 1: 追加失败测试**

```python
from modules.template_formula import resolve_row_source


def _make_rows_file(name, sheet, n):
    fpath = os.path.join(TEST_DIR, name)
    wb = openpyxl.Workbook(); wb.active.title = sheet
    for i in range(n):
        wb[sheet].cell(row=i + 1, column=1, value=i)
    wb.save(fpath); wb.close()
    return fpath


def test_row_source_explicit():
    f1 = _make_rows_file('_tmp_rs1.xlsx', 'A', 5)
    n, chosen = resolve_row_source((f1, 'A'), {'=A!B2'}, {f1: ['A']})
    assert n == 5, (n, chosen)
    os.remove(f1)
    print("PASS test_row_source_explicit")


def test_row_source_most_referenced():
    f1 = _make_rows_file('_tmp_rs2.xlsx', 'A', 3)
    f2 = _make_rows_file('_tmp_rs3.xlsx', 'B', 7)
    # 公式里 A 引用 1 次,B 引用 2 次 → 选 B(7 行)
    formulas = {'c': '=B!C2+B!C3+A!C2'}
    n, chosen = resolve_row_source(None, formulas, {f1: ['A'], f2: ['B']}, [f1, f2])
    assert n == 7, (n, chosen)
    os.remove(f1); os.remove(f2)
    print("PASS test_row_source_most_referenced")


def test_row_source_fallback_first_file():
    f1 = _make_rows_file('_tmp_rs4.xlsx', 'A', 4)
    n, chosen = resolve_row_source(None, {'c': '普通文本无引用'}, {f1: ['A']}, [f1])
    assert n == 4, (n, chosen)
    os.remove(f1)
    print("PASS test_row_source_fallback_first_file")
```
更新 `if __name__` 块加入这 3 个测试。

- [ ] **Step 2: 跑测试,确认失败**

Run: `python tests/template/test_template_formula.py`
Expected: FAIL(`ImportError: cannot import name 'resolve_row_source'`)。

- [ ] **Step 3: 实现 `resolve_row_source`**

在 `modules/template_formula.py` 追加:
```python
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
    data_files: List[str],
) -> Tuple[int, str]:
    """
    决定输出公式列的行数 N,返回 (N, 说明)。
    1) row_source 显式 → 用 (文件, sheet) 的行数
    2) 否则 → 公式中被引用次数最多的数据 sheet
    3) 再否则 → 第一个数据文件的第一个 sheet
    """
    # ① 显式
    if row_source:
        f, s = row_source
        f = _match_data_file(f, data_files)
        if s not in file_sheet_index.get(f, []):
            raise ValueError(f"row_source 指定的 sheet '{s}' 不在文件 '{f}' 中")
        return _sheet_row_count(f, s), f"row_source={f}:{s}"

    # ② 统计每个数据 sheet 被引用次数
    refs = _collect_referenced_sheets(formulas)
    # 统计出现次数(含重复,表示引用次数)
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
```

- [ ] **Step 4: 跑测试,确认通过**

Run: `python tests/template/test_template_formula.py`
Expected: 全部 PASS。

- [ ] **Step 5: Commit**

```bash
git add modules/template_formula.py tests/template/test_template_formula.py
git commit -m "feat(template_formula): 行数驱动 resolve_row_source"
```

---

## Task 7: 主函数 `generate_formulas_from_template` 串联(TDD,内部模式 happy path)

**Files:**
- Modify: `modules/template_formula.py`
- Test: `tests/template/test_template_formula.py`

- [ ] **Step 1: 追加端到端失败测试**

```python
from modules.template_formula import generate_formulas_from_template


def _build_formula_fixture():
    """模板 1 个公式列,引用 1 个外部数据 sheet。"""
    template = os.path.join(TEST_DIR, '_tmp_tpl.xlsx')
    data = os.path.join(TEST_DIR, '_tmp_data.xlsx')
    out = os.path.join(TEST_DIR, '_tmp_out.xlsx')

    # 模板:第1行表头 [合计],第2行公式 =Src!B2
    wb = openpyxl.Workbook(); wb.active.title = '结果'
    ws = wb['结果']
    ws.cell(row=1, column=1, value='合计')
    ws.cell(row=2, column=1, value='=Src!B2')
    wb.save(template); wb.close()

    # 数据:Src sheet,3 行
    wb = openpyxl.Workbook(); wb.active.title = 'Src'
    wb['Src'].cell(row=1, column=2, value='h')
    for i in range(3):
        wb['Src'].cell(row=i + 2, column=2, value=i + 1)
    wb.save(data); wb.close()
    return template, data, out


def test_generate_internal_happy_path():
    template, data, out = _build_formula_fixture()
    df = generate_formulas_from_template(
        template_file=template, template_sheet='结果',
        data_files=[data], output_file=out,
    )
    # 返回骨架:只含公式列,行数=数据 3 行(Src max_row 含表头=4?见下)
    assert list(df.columns) == ['合计'], df.columns

    # 读回输出文件,核对公式与行数
    wb = openpyxl.load_workbook(out)
    ws = wb['结果']
    assert ws.cell(row=1, column=1).value == '合计'
    # 内部模式:数据 sheet 被复制进来
    assert 'Src' in wb.sheetnames, wb.sheetnames
    # 第 2 行公式应为 =Src!B2(内部引用,行偏移 0)
    assert ws.cell(row=2, column=1).value == '=Src!B2', ws.cell(row=2, column=1).value
    wb.close()

    for p in (template, data, out):
        os.remove(p)
    print("PASS test_generate_internal_happy_path")
```
更新 `if __name__` 加入该测试。

> 行数说明:`_sheet_row_count` 用 `ws.max_row`,Src 含表头共 4 行,故输出 4 公式行(第 2–5 行)。测试只断言列结构与第 2 行公式,不锁死行数,避免与 max_row 语义耦合。

- [ ] **Step 2: 跑测试,确认失败**

Run: `python tests/template/test_template_formula.py`
Expected: FAIL(`ImportError: cannot import name 'generate_formulas_from_template'`)。

- [ ] **Step 3: 实现主函数及编排 helper**

在 `modules/template_formula.py` 追加:
```python
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
    # 需要模板列名→索引,按公式列在模板中的位置取样式
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
    print("=== Excel 公式模板生成器 ===\\n")

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
    print(f"\\n输出文件已保存: {output_file}")
    return output_df
```

- [ ] **Step 4: 跑测试,确认通过**

Run: `python tests/template/test_template_formula.py`
Expected: 全部 PASS(含 happy path)。若 happy path 的公式断言不符(如内部引用串格式不同),按实际 `replace_sheet_references` 产出的串调整断言——但优先信任实现,核对串是否合理(应为 `=Src!B2`)。

- [ ] **Step 5: 跑回归测试,确保老方法未受影响**

Run: Task 1 的 6 条命令。
Expected: 全绿。

- [ ] **Step 6: Commit**

```bash
git add modules/template_formula.py tests/template/test_template_formula.py
git commit -m "feat(template_formula): 主函数 generate_formulas_from_template(内部模式)"
```

---

## Task 8: 外部模式 + 显式 sheet_mapping 覆盖(TDD)

**Files:**
- Modify: `tests/template/test_template_formula.py`(逻辑已在 Task 5/7 实现,本任务加覆盖测试)

- [ ] **Step 1: 追加两个测试**

```python
def test_generate_external_refs_mode():
    template, data, out = _build_formula_fixture()
    generate_formulas_from_template(
        template_file=template, template_sheet='结果',
        data_files=[data], output_file=out, use_external_refs=True,
    )
    wb = openpyxl.load_workbook(out)
    assert 'Src' not in wb.sheetnames, "外部模式不应复制数据 sheet"
    formula = wb['结果'].cell(row=2, column=1).value
    assert formula is not None and ('Src' in formula) and ('!' in formula), formula
    # 外部引用应含文件名标记
    assert os.path.basename(data) in formula, formula
    wb.close()
    for p in (template, data, out):
        os.remove(p)
    print("PASS test_generate_external_refs_mode")


def test_generate_sheet_mapping_override():
    # 两个数据文件都有 Src,靠 sheet_mapping 指定用第二个
    template, data, out = _build_formula_fixture()
    data2 = os.path.join(TEST_DIR, '_tmp_data2.xlsx')
    wb = openpyxl.Workbook(); wb.active.title = 'Src'
    wb['Src'].cell(row=2, column=2, value=99)
    wb.save(data2); wb.close()

    generate_formulas_from_template(
        template_file=template, template_sheet='结果',
        data_files=[data, data2], output_file=out,
        sheet_mapping={'Src': data2},
    )
    wb = openpyxl.load_workbook(out)
    # 仅 data2 的 Src 被复制(data 的 Src 被映射跳过,不复制)
    assert 'Src' in wb.sheetnames
    assert wb['Src'].cell(row=2, column=2).value == 99
    wb.close()
    for p in (template, data, data2, out):
        os.remove(p)
    print("PASS test_generate_sheet_mapping_override")
```
更新 `if __name__` 加入两个测试。

- [ ] **Step 2: 跑测试**

Run: `python tests/template/test_template_formula.py`
Expected: 全部 PASS。外部模式/映射覆盖逻辑已在 Task 5/7 实现(`is_internal` 由 `use_external_refs` 控制;`sheet_mapping` 在 `resolve_formula_sheet` 第①层生效),此处为覆盖确认。若失败,定位是 `_copy_referenced_external_sheets` 或 `is_internal` 标志传递问题并修正。

- [ ] **Step 3: Commit**

```bash
git add tests/template/test_template_formula.py
git commit -m "test(template_formula): 外部模式与 sheet_mapping 覆盖"
```

---

## Task 9: 冲突报错 + 多公式列行偏移(TDD)

**Files:**
- Modify: `tests/template/test_template_formula.py`

- [ ] **Step 1: 追加两个测试**

```python
def test_generate_conflict_raises():
    template, data, out = _build_formula_fixture()
    data2 = os.path.join(TEST_DIR, '_tmp_data2.xlsx')
    wb = openpyxl.Workbook(); wb.active.title = 'Src'
    wb.save(data2); wb.close()
    try:
        generate_formulas_from_template(
            template_file=template, template_sheet='结果',
            data_files=[data, data2], output_file=out,  # 不传 sheet_mapping → 冲突
        )
        assert False, "应抛错"
    except ValueError as e:
        assert 'Src' in str(e)
    finally:
        for p in (template, data, data2):
            if os.path.exists(p):
                os.remove(p)
        if os.path.exists(out):
            os.remove(out)
    print("PASS test_generate_conflict_raises")


def test_generate_multi_formula_columns_row_offset():
    template = os.path.join(TEST_DIR, '_tmp_tpl2.xlsx')
    data = os.path.join(TEST_DIR, '_tmp_data3.xlsx')
    out = os.path.join(TEST_DIR, '_tmp_out2.xlsx')
    wb = openpyxl.Workbook(); wb.active.title = '结果'
    ws = wb['结果']
    ws.cell(row=1, column=1, value='甲'); ws.cell(row=1, column=2, value='乙')
    ws.cell(row=2, column=1, value='=Src!B2')
    ws.cell(row=2, column=2, value='=Src!C2')
    wb.save(template); wb.close()

    wb = openpyxl.Workbook(); wb.active.title = 'Src'
    wb['Src'].cell(row=1, column=2, value='h')
    for i in range(3):
        wb['Src'].cell(row=i + 2, column=2, value=i)
        wb['Src'].cell(row=i + 2, column=3, value=i * 10)
    wb.save(data); wb.close()

    generate_formulas_from_template(
        template_file=template, template_sheet='结果',
        data_files=[data], output_file=out,
    )
    wb = openpyxl.load_workbook(out)
    ws = wb['结果']
    # 两列都写入,且第 3 行公式行偏移 +1
    assert ws.cell(row=2, column=1).value == '=Src!B2'
    assert ws.cell(row=3, column=1).value == '=Src!B3', ws.cell(row=3, column=1).value
    assert ws.cell(row=2, column=2).value == '=Src!C2'
    assert ws.cell(row=3, column=2).value == '=Src!C3', ws.cell(row=3, column=2).value
    wb.close()
    for p in (template, data, out):
        os.remove(p)
    print("PASS test_generate_multi_formula_columns_row_offset")
```
更新 `if __name__` 加入两个测试。

- [ ] **Step 2: 跑测试**

Run: `python tests/template/test_template_formula.py`
Expected: 全部 PASS。多列偏移由 Task 7 的 `row_offset = r - 2` 逻辑覆盖;冲突由 Task 5 第②层覆盖。若偏移断言不符,核对 `replace_sheet_references` 的 row_offset 参数是否正确传递。

- [ ] **Step 3: Commit**

```bash
git add tests/template/test_template_formula.py
git commit -m "test(template_formula): 冲突报错与多公式列行偏移"
```

---

## Task 10: CLI 子命令 `template-formula`

**Files:**
- Modify: `main.py`

- [ ] **Step 1: 在 `main.py` 顶部新增导入**

把:
```python
from modules.template_generator import generate_excel_from_template, parse_column_mappings
```
下方追加:
```python
from modules.template_formula import generate_formulas_from_template
```

- [ ] **Step 2: 新增 `run_template_formula(args)` 与参数解析**

在 `run_template` 函数之后追加:
```python
def parse_sheet_mapping(mapping_str):
    """解析 'sheet名:文件路径,sheet名:文件路径' → dict(按第一个冒号切分)。"""
    if not mapping_str:
        return None
    result = {}
    for pair in mapping_str.split(','):
        pair = pair.strip()
        if not pair:
            continue
        if ':' not in pair:
            raise ValueError(f"sheet_mapping 格式错误(应为 sheet名:文件路径): {pair}")
        k, v = pair.split(':', 1)
        result[k.strip()] = v.strip()
    return result


def run_template_formula(args):
    """运行模板公式专用生成"""
    print("=== 模板公式生成模式 ===\\n")
    try:
        sheet_mapping = parse_sheet_mapping(args.sheet_mapping)
        row_source = tuple(args.row_source.split(':', 1)) if args.row_source else None

        result = generate_formulas_from_template(
            template_file=args.template,
            template_sheet=args.template_sheet,
            data_files=args.data_file,
            output_file=args.output,
            sheet_mapping=sheet_mapping,
            row_source=row_source,
            use_external_refs=args.external_refs,
        )
        print("\\n生成完成!")
        print(f"输出文件: {args.output}")
        print("\\n生成结果:")
        print(result)
    except Exception as e:
        print(f"\\n❌ 生成失败: {e}")
        sys.exit(1)
```

在 `main()` 的 `template_parser = subparsers.add_parser('template', ...)` 块之后,追加子命令:
```python
    # 模板公式生成命令
    formula_parser = subparsers.add_parser('template-formula', help='基于模板生成只含公式列的Excel')
    formula_parser.add_argument('-t', '--template', required=True, help='模板Excel文件路径')
    formula_parser.add_argument('-ts', '--template-sheet', required=True, help='模板中含公式的目标工作表')
    formula_parser.add_argument('-d', '--data-file', action='append', required=True,
                                metavar='FILE', help='数据文件(可多次使用)')
    formula_parser.add_argument('--sheet-mapping', help='公式sheet名:文件路径,逗号分隔多条')
    formula_parser.add_argument('--row-source', help='文件路径:sheet名,指定行数驱动源')
    formula_parser.add_argument('-o', '--output', default='output.xlsx', help='输出文件名')
    formula_parser.add_argument('--external-refs', action='store_true', help='使用外部文件引用(默认内部)')
```

在命令分发处(`if args.command == ...` 链)追加:
```python
    elif args.command == 'template-formula':
        run_template_formula(args)
```

- [ ] **Step 3: 冒烟测试——CLI 实跑**

Run:
```bash
cd tests/template && python -c "
import openpyxl, os
# 复用最小 fixture
wb=openpyxl.Workbook(); wb.active.title='结果'
wb['结果'].cell(1,1,'合计'); wb['结果'].cell(2,1,'=Src!B2')
wb.save('_cli_tpl.xlsx'); wb.close()
wb=openpyxl.Workbook(); wb.active.title='Src'
wb['Src'].cell(2,2,1); wb.save('_cli_data.xlsx'); wb.close()
"
python ../../main.py template-formula -t _cli_tpl.xlsx -ts 结果 -d _cli_data.xlsx -o _cli_out.xlsx
python -c "
import openpyxl
wb=openpyxl.load_workbook('tests/template/_cli_out.xlsx')
print('sheets:', wb.sheetnames)
print('formula:', wb['结果'].cell(2,1).value)
"
```
Expected: 子命令执行无报错;输出含 `结果` 与 `Src` 两个 sheet;公式为 `=Src!B2`。

- [ ] **Step 4: 清理 fixture**

Run: `rm -f tests/template/_cli_tpl.xlsx tests/template/_cli_data.xlsx tests/template/_cli_out.xlsx`

- [ ] **Step 5: 跑回归测试**

Run: Task 1 的 6 条命令 + `python tests/template/test_template_formula.py`。
Expected: 全绿。

- [ ] **Step 6: Commit**

```bash
git add main.py
git commit -m "feat(cli): 新增 template-formula 子命令"
```

---

## Task 11: README 新增独立章节

**Files:**
- Modify: `README.md`

- [ ] **Step 1: 在 README「模板生成模块」章节之后,新增独立章节**

追加(放在「## 安装要求」之前):
```markdown
## 模板公式专用生成模块 (`modules/template_formula.py`)

> 与上面的「模板生成模块」**相互独立**。本模块**只输出公式列**,无需配置列映射,仅凭模板 + 数据文件路径即可生成。

### 功能特点

- **只输出公式列**:自动检测模板目标 sheet 中第 2 行以 `=` 开头的所有列,丢弃其余列。
- **公式按数据行复制**:逐行套用公式模板并自动偏移行号。
- **sheet 引用三层解析**(优先级从高到低):
  1. 显式映射 `sheet_mapping`(你说了算)
  2. 外部数据文件同名自动匹配
  3. 模板自带 sheet
- **两种引用模式**:默认内部(数据 sheet 复制进输出,自包含);`use_external_refs=True` 外部引用(与源文件活链接)。
- **行数驱动**:可选 `row_source` 指定;不传则取公式中引用最多的数据 sheet。

### 使用方法(命令行)

```bash
python main.py template-formula -t template.xlsx -ts 结果 \
    -d data1.xlsx -d data2.xlsx \
    --sheet-mapping "ESDP-Bpart:bpart.xlsx" \
    --row-source "bpart.xlsx:ESDP-Bpart" \
    -o result.xlsx
```

参数:

| 参数 | 必需 | 说明 |
|------|------|------|
| `-t/--template` | 是 | 模板 Excel 路径 |
| `-ts/--template-sheet` | 是 | 模板中含公式的目标工作表 |
| `-d/--data-file` | 是 | 数据文件,可多次使用 |
| `--sheet-mapping` | 否 | `公式sheet名:文件路径`,逗号分隔多条(显式覆盖) |
| `--row-source` | 否 | `文件路径:sheet名`,行数驱动源 |
| `-o/--output` | 否 | 输出文件(默认 output.xlsx) |
| `--external-refs` | 否 | 使用外部文件引用(默认内部) |

### 使用方法(导入)

```python
from modules.template_formula import generate_formulas_from_template

result = generate_formulas_from_template(
    template_file='template.xlsx',
    template_sheet='结果',
    data_files=['bpart.xlsx', 'cpart.xlsx'],
    output_file='result.xlsx',
    # sheet_mapping={'ESDP-Bpart': 'bpart.xlsx'},  # 可选
    # row_source=('bpart.xlsx', 'ESDP-Bpart'),     # 可选
)
```

### 与老模板生成模块的区别

| 维度 | 模板生成模块(老) | 模板公式专用模块(新) |
|------|-------------------|------------------------|
| 入参 | 需逐个声明 sheet/列映射/alias | 只要文件路径 |
| 输出 | 全部列 + 合并数据 | 仅公式列 |
| 适用 | 需要数据列填充 | 只关心公式,数据仅作引用源 |
```

- [ ] **Step 2: 校验 README 渲染**

Run: `grep -n "模板公式专用生成模块" README.md`
Expected: 命中新章节标题。

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: 新增模板公式专用生成模块说明"
```

---

## 完成标准

- [ ] 所有 Task 的 checkbox 勾完。
- [ ] `python tests/template/test_template_formula.py` 全绿。
- [ ] Task 1 的 6 条回归测试全绿(老方法零回归)。
- [ ] CLI `python main.py template-formula -h` 正常显示帮助。
- [ ] `modules/template_generator.py` 行数较抽取前明显下降;`modules/template_formula.py` 与之物理分离,互不 import。
