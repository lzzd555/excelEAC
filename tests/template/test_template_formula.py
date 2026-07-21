"""
测试 generate_formulas_from_template(模板公式专用生成)
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import openpyxl
from modules.template_formula import discover_file_sheets, resolve_formula_sheet, resolve_row_source

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


def _make_rows_file(name, sheet, n):
    fpath = os.path.join(TEST_DIR, name)
    wb = openpyxl.Workbook(); wb.active.title = sheet
    for i in range(n):
        wb[sheet].cell(row=i + 1, column=1, value=i)
    wb.save(fpath); wb.close()
    return fpath


def test_row_source_explicit():
    f1 = _make_rows_file('_tmp_rs1.xlsx', 'A', 5)
    n, chosen = resolve_row_source((f1, 'A'), {'c': '=A!B2'}, {f1: ['A']})
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


if __name__ == '__main__':
    test_discover_file_sheets_lists_all()
    test_resolve_explicit_mapping_wins()
    test_resolve_auto_match_single()
    test_resolve_conflict_without_mapping_raises()
    test_resolve_template_self_sheet()
    test_resolve_not_found_raises()
    test_row_source_explicit()
    test_row_source_most_referenced()
    test_row_source_fallback_first_file()
    print("all pass")
