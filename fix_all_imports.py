#!/usr/bin/env python3
"""修复测试文件的导入路径问题"""
import os

def fix_validation_imports():
    # 修复validation测试文件的导入
    validation_files = [
        'tests/validation/test_abnormal_detail.py',
        'tests/validation/test_standard.py',
        'tests/validation/test_string_columns.py',
    ]

    for file_path in validation_files:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 替换旧的导入为新的导入
        old_import = 'from excel_validator import process_excel_with_validation'
        new_import = '''import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation'''

        if old_import in content:
            content = content.replace(old_import, new_import)
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f'✅ 已修复: {file_path}')
        else:
            print(f'⏭️  无需修复: {file_path}')

def fix_merge_imports():
    # 修复merge测试文件的导入
    merge_files = [
        'tests/merge/test_basic_merge.py',
        'tests/merge/test_extra_columns.py',
        'tests/merge/test_multi_column.py',
        'tests/merge/test_merge.py',
    ]

    for file_path in merge_files:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 替换旧的导入为新的导入
        old_import = 'from modules.merge import merge_excel_tables'
        new_import = '''import pandas as pd
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables'''

        if old_import in content:
            content = content.replace(old_import, new_import)
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f'✅ 已修复: {file_path}')
        else:
            print(f'⏭️  无需修复: {file_path}')

if __name__ == '__main__':
    print('=== 修复测试文件导入路径 ===\n')
    fix_validation_imports()
    print()
    fix_merge_imports()
    print('\n✅ 所有导入路径已修复')
