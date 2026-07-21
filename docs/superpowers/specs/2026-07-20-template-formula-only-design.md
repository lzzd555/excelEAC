# 模板公式专用生成功能 设计文档

- 日期: 2026-07-20
- 状态: 已确认,待实现
- 相关模块: `modules/template_generator.py`(老)、新增 `modules/template_formula.py`、新增 `modules/_template_core.py`

## 1. 背景与目标

现有 `generate_excel_from_template` 把「列映射 + 数据合并 + 公式改写」耦合在一个流程里,入参重(`column_mappings`、`alias`、`sheet_name` 都要逐个声明)。

目标场景:用户手头有一个含公式的模板 + 若干数据 Excel,只想**快速生成只含公式列的结果**,不想配置列映射。

**目标**:提供一个与老方法清晰区分的新入口,仅凭「模板 + 数据文件路径」即可生成只含公式列的输出;公式里的 sheet 引用按「显式映射 > 外部数据文件 > 模板自带 sheet」三层优先级自动解析。

**非目标**:
- 不做列映射、不做数据列填充(模板中非公式列全部丢弃)。
- 不改动老方法 `generate_excel_from_template` 的行为与接口。

## 2. 用户面分离(硬约束)

为降低用户学习压力,新老方法在用户面完全分开,禁止任何形式的「同一函数加 mode 开关」兼底:

| 维度 | 老方法 | 新方法 |
|------|--------|--------|
| 函数 | `generate_excel_from_template` | `generate_formulas_from_template` |
| CLI 子命令 | `template` | `template-formula` |
| README | 现有章节 | 新增独立章节 |
| 内部依赖 | `template_generator.py` | `template_formula.py` |
| 共享底层 | 均从 `_template_core.py` 导入 | 均从 `_template_core.py` 导入 |

## 3. 文件结构(方案 C:抽共享核心)

```
modules/
├── _template_core.py        新建:纯共享原语(内部,_ 前缀表示非用户 API)
├── template_generator.py    老方法瘦身:列映射 + 数据合并 + 老流程编排
├── template_formula.py      新方法:只出公式列的编排 + 3 个新 helper
├── validation.py
└── merge.py
```

**抽取原则**:只搬「给一个 cell/formula 就能干活、不依赖输出布局」的纯函数。带布局假设的(如老方法按整列刷样式的 `apply_template_styles`)留在各自方法文件。

## 4. 抽取范围

### 移入 `_template_core.py`(共享纯原语)

- **数据类**:`CellStyle`
- **公式引擎**:`parse_formula_references`、`replace_sheet_references`、`_adjust_cell_ref`、`_adjust_single_ref`、`_adjust_range_ref`、`_extract_sheet_name`、`_find_matching_info`、`_resolve_multiple_matches`、`_find_by_index`、`_build_reference`、`_replace_quoted_match`、`_replace_unquoted_match`、`_replace_bracket_match`
- **外部链接**:`read_external_links`、`_extract_external_links`、`_parse_single_link`、`replace_link_indices_with_filenames`
- **模板读取**:`read_template_structure`、`_read_template_columns`、`_read_formula_templates`
- **样式 / 颜色复制**:`copy_color` 及全部 `_copy_rgb/_theme/_indexed_color`、`copy_side`、`copy_cell_style`、`_copy_font_style`、`_get_font_color`、`_copy_fill_style`、`_apply_fill_by_type`、`_apply_solid_fill`、`_get_fill_color_value`、`_apply_fallback_fill`、`_copy_border_style`、`_copy_alignment_style`、`_copy_worksheet`

### 留在 `template_generator.py`(老方法专属)

- 列映射 / 数据合并:`read_data_source`、`_apply_column_mappings`、`_process_string_columns`、`merge_data_by_row`
- 老方法 alias 构建:`_init_alias_to_info`、`_create_data_source_config`、`_update_alias_mapping`
- 老流程编排:`_validate_input_files`、`_analyze_template`、`_load_all_data_sources`、`_generate_output_file`、`_copy_data_source_sheets`、`_prepare_output_df`、`_apply_formulas`、`apply_formulas_to_output`、`apply_template_styles` 及其子 helper、`_filter_data_by_primary_column`、`_apply_string_column_format`、`_print_formula_summary`、`_print_data_source_mapping`、`parse_column_mappings`、`generate_excel_from_template`
- 老方法专属数据类:`DataColumnMapping`、`DataSourceConfig`

### `template_formula.py`(新方法专属)

- 新增 helper:`discover_file_sheets`、`build_alias_to_info_from_files`、`resolve_row_source`
- 新流程编排 + `generate_formulas_from_template`

## 5. 新函数签名

```python
def generate_formulas_from_template(
    template_file: str,
    template_sheet: str,                            # 模板中含公式的目标工作表(必需)
    data_files: List[str],                          # 仅文件路径,自动发现所有 sheet
    output_file: str,
    sheet_mapping: Optional[Dict[str, str]] = None, # 显式映射:公式 sheet 名 → 数据文件路径
    row_source: Optional[Tuple[str, str]] = None,   # (文件路径, sheet 名) 驱动行数;不传则自动
    use_external_refs: bool = False,                # False=内部(默认),True=外部引用
) -> pd.DataFrame
```

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `template_file` | str | 是 | 模板 Excel 路径 |
| `template_sheet` | str | 是 | 模板中含公式的目标工作表(输出的列结构以此为准) |
| `data_files` | str 列表 | 是 | 数据 Excel 路径列表(只要路径,自动发现其中所有 sheet) |
| `output_file` | str | 是 | 输出路径(相对路径自动转绝对) |
| `sheet_mapping` | dict | 否 | 公式 sheet 名 → 数据文件路径,显式声明,最高优先级 |
| `row_source` | (str, str) | 否 | 驱动行数的 (文件, sheet);不传则自动 |
| `use_external_refs` | bool | 否 | 默认 False 内部模式;True 外部引用模式 |

> **公式列无需指定**:自动检测 `template_sheet` 中第 2 行以 `=` 开头的**所有**列作为公式列,全部输出。

**返回**:`pd.DataFrame`,只含公式列、N 行的数据骨架(单元格值为空,公式实际写入 Excel 文件而非 DataFrame),供调用方检查列结构与行数。

## 6. 核心数据流

1. **校验**:模板与各 `data_files` 存在。
2. **分析模板**:读列名(第 1 行)、读公式(自动检测第 2 行以 `=` 开头的**所有**列作为公式列)、取 `template_ws` 备样式、读外部链接。
3. **发现 sheet**:对每个 `data_files` 调 `discover_file_sheets` → 全部 `(file, sheet)` 清单。
4. **建 `alias_to_info`**:`build_alias_to_info_from_files` 按第 7 节三层优先级构建。
5. **定行数**:`resolve_row_source` 确定 N(第 8 节)。
6. **构造输出**:只含公式列、N 行的 DataFrame。
7. **写出**:`pd.ExcelWriter` → 写 `结果` sheet →(内部模式)复制被引用的数据/模板 sheet → 对公式列逐行套用改写后的公式(行偏移)→ 用核心 `copy_cell_style` 给公式列贴样式。
8. **返回** 公式列 DataFrame。

## 7. sheet 引用三层解析(核心规则)

对公式里每个 sheet 引用(如 `数据源!C:C` 中的「数据源」),按优先级查找:

1. **① `sheet_mapping`(显式映射)**:若该 sheet 名是 `sheet_mapping` 的键 → 用声明的文件,**停止**。
2. **② 外部数据文件(`data_files`)**:扫描所有数据文件,若有唯一同名 sheet → 用之,**停止**。
3. **③ 模板文件自身**:若模板 Excel 含同名 sheet → 当模板内置数据源,**停止**。
4. **三层都找不到** → 抛错,信息含「公式列 X 引用 sheet 'Y',映射/数据文件/模板中均未找到,可用 `--sheet-mapping` 指定」。

**冲突处理**:第②层出现同名 sheet 跨多个文件、且未在第①层显式消歧 → 抛错,列出冲突文件,要求用 `--sheet_mapping` 指明。

**模板自引用**:模板的 `template_sheet` 引用自己 → 指向输出的 `结果` sheet(沿用现有 `is_template_self_reference` 语义)。

## 8. 行数驱动

输出公式列的行数 N 由以下顺序决定,每次都打印实际选择:

1. 传了 `row_source` → 用该 (文件, sheet) 的行数。
2. 否则 → 取所有公式里被引用次数最多的数据 sheet 的行数(并列时按 `data_files` 顺序取第一个)。
3. 再否则 → 第一个数据文件的第一个 sheet 的行数。

## 9. 内部 / 外部模式

沿用老方法语义:

- **内部模式(默认 `use_external_refs=False`)**:仅把公式实际引用到的那些 sheet(无论来自外部数据文件还是模板自带)复制进输出文件,公式指向内部副本。产物自包含,数据为生成时快照。
- **外部模式(`use_external_refs=True`)**:不复制,公式直接引用外部文件(`'[file.xlsx]Sheet'!A1`)。输出小、与源文件保持活链接,但打开时需能访问源文件。

## 10. 错误处理

| 场景 | 行为 |
|------|------|
| 模板 / 数据文件不存在 | `FileNotFoundError`(沿用现有) |
| 公式 sheet 引用三层未解析 | 抛错,提示用 `--sheet-mapping` |
| 同名 sheet 跨文件冲突未消歧 | 抛错,列出冲突文件 |
| `row_source` 指定 sheet 不存在 | 抛错 |
| 模板中无任何公式列(第 2 行无 `=` 单元格) | 抛错提示「未检测到公式列」 |

## 11. CLI

新增 `template-formula` 子命令:

```bash
python main.py template-formula \
    -t template.xlsx -ts Sheet1 \
    -d data1.xlsx -d data2.xlsx \
    --sheet-mapping "ESDP-Bpart:bpart.xlsx,ESDP-Cpart:cpart.xlsx" \
    --row-source "bpart.xlsx:ESDP-Bpart" \
    -o result.xlsx
```

参数:

| 参数 | 说明 |
|------|------|
| `-t/--template` | 模板路径(必需) |
| `-ts/--template-sheet` | 模板中含公式的目标工作表(必需) |
| `-d/--data-file` | 数据文件,可多次使用(必需,至少一个) |
| `--sheet-mapping` | `公式sheet名:文件路径`,逗号分隔多条(可选) |
| `--row-source` | `文件路径:sheet名`(可选) |
| `-o/--output` | 输出文件(默认 output.xlsx) |
| `--external-refs` | 切外部引用模式(可选) |

## 12. 测试

放在 `tests/template/`:

**单元**:
- `discover_file_sheets`:正确列出文件全部 sheet。
- `build_alias_to_info_from_files`:① 显式映射优先;② 自动同名匹配;③ 模板自带 sheet;④ 同名冲突抛错。
- `resolve_row_source`:`row_source` 显式 / 最常引用 / 兜底第一个文件,三种路径。

**集成**:
- 自动按名匹配成功生成,公式行偏移正确。
- 显式 `sheet_mapping` 覆盖自动结果。
- 同名冲突按预期报错。
- 多公式列逐行偏移正确。
- 模板自带 sheet 被解析(第③层)。
- 内部 vs 外部两种模式产物均能被 openpyxl 正确读回、公式串符合预期。

## 13. 实现风险与兜底

- 搬到 `_template_core.py` 的都是纯原语,老方法改为从核心导入,行为零变化。
- 抽取后**先跑现有 `tests/` 全套**确认老方法无回归,再加新功能测试。
- `template_generator.py` 从 1432 行明显瘦身,顺带缓解文件过大的历史问题。
- 行数驱动有自动兜底,但默认行为会 log,避免黑盒。
