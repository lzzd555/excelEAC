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
