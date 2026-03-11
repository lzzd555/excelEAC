#!/usr/bin/env python3
import os
import re

def fix_validation_test_imports():
    files = [
        'tests/validation/test_abnormal_detail.py',
        'tests/validation/test_standard.py',
        'tests/validation/test_string_columns.py',
    ]
    for file_path in files:
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            continue
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        old_pattern = r'from excel_validator import process_excel_with_validation'
        new_import = '''import pandas as pd
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation'''
        content = re.sub(old_pattern, new_import, content)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"已修复: {file_path}")
if __name__ == '__main__':
    fix_validation_test_imports()
