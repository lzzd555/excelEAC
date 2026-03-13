# 模板自引用功能测试总结

## 问题背景

在模板生成模块中，如果模板中显式引用了模板自身所在表格的单元格，最终生成的文件中使用的公式会输出 `[2]xxx.xlsx` 这样有问题的文件引用，而不是正确的本地引用。

## 解决方案

修改 `modules/template_generator.py` 中的以下函数：

1. **`replace_sheet_references` 函数**：
   - 添加 `output_file_path` 参数
   - 在替换函数中检测是否是本地引用
   - 使用 `is_template_self_reference` 标记识别模板自引用

2. **`apply_formulas_to_output` 函数**：
   - 添加 `output_file_path` 参数并传递给 `replace_sheet_references`

3. **`generate_excel_from_template` 函数**：
   - 添加模板 sheet 名称到输出 sheet 的映射
   - 使用 `is_template_self_reference` 标记
   - 在调用时传递 `output_file_path` 参数

## 测试结果

### 测试 1: 基础自引用测试

**模板文件**: `template_with_self_ref.xlsx`
- Sheet 名称: `TemplateSheet`
- 公式格式: `=A2-B2-C2` (纯本地引用)

**输出文件**: `test_self_ref_output.xlsx`
- 公式格式: `=A2-B2-C2`
- 结果: ✅ 通过

### 测试 2: 带 Sheet 名称的自引用测试

**模板文件**: `template_with_sheet_ref.xlsx`
- Sheet 名称: `MyTemplateSheet`
- 公式格式: `=MyTemplateSheet!A2-MyTemplateSheet!B2` (带 sheet 名称)

**输出文件**: `test_sheet_ref_output.xlsx`
- 公式格式: `=结果!A2-结果!B2`
- 结果: ✅ 通过
- 说明: Sheet 名称 `MyTemplateSheet` 正确映射到 `结果`

### 测试 3: Excel 索引格式测试

**模板文件**: `template_with_index_ref.xlsx`
- Sheet 名称: `MyTemplateSheet`
- 公式格式: `=[2]MyTemplateSheet!A2` (Excel 外部引用索引)

**输出文件**: `test_index_ref_output.xlsx`
- 公式格式: `='结果'!A2`
- 结果: ✅ 通过
- 说明: Excel 索引 `[2]` 正确映射为本地引用 `结果!A2`

### 测试 4: 混合引用测试

**模板文件**: `template_mixed_refs.xlsx`
- Sheet 名称: `MyTemplateSheet`
- 包含自引用和外部引用:
  - 自引用: `=A2-B2-C2`
  - 外部引用: `=[1]DataSheet!D2`

**输出文件**: `test_mixed_refs_output.xlsx`
- 自引用公式: `=A2-B2-C2` (本地引用)
- 外部引用公式: `='[external_data.xlsx]DataSheet'!D2`
- 结果: ✅ 通过
- 说明: 混合引用处理正确

## 关键改进

### 1. 本地引用检测

```python
def is_local_reference(info: Dict[str, str]) -> bool:
    """检测引用是否是输出文件自身的本地引用"""
    return (
        (output_file_path and os.path.normpath(file_path) == os.path.normpath(output_file_path)) or
        info.get('is_template_self_reference', False)
    )
```

### 2. 模板 Sheet 映射

在 `generate_excel_from_template` 中添加：

```python
template_sheet_info = {
    'file_path': output_file,
    'sheet_name': '结果',
    'is_template_self_reference': True
}
alias_to_info[template_sheet.lower()] = template_sheet_info
```

### 3. 本地引用格式

当检测到本地引用时，返回本地引用格式：
- 带引号: `'结果'!A2`
- 不带引号: `结果!A2` (如果 sheet 名不需要引号)

当检测到外部引用时，返回外部引用格式：
- `'[filename.xlsx]SheetName'!A2`

## 测试文件列表

- `test_template_self_ref.py` - 基础自引用测试
- `test_template_self_ref_advanced.py` - 带 sheet 名称的自引用测试
- `test_excel_index_ref.py` - Excel 索引格式测试
- `test_mixed_references.py` - 混合引用测试

## 结论

✅ 所有测试通过
✅ 模板自引用正确映射到本地引用
✅ 外部引用正确保持文件名格式
✅ 混合引用场景处理正确

修改后的代码能够正确处理模板中引用自身工作表的情况，避免了生成无效的外部引用格式。
