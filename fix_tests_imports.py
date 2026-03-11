#!/usr/bin/env python3
"""批量修复所有测试文件的import路径问题"""

import os

def fix_file(file_path):
    """修复单个文件的import路径"""
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 查找所有import语句
    lines = content.split('\n')
    new_lines = []
    skip_until_next = False

    for line in lines:
        # 跳过旧的导入
        if 'from excel_validator import' in line:
            skip_until_next = True
            continue

        # 在第一个函数之前添加正确的导入
        if skip_until_next and line.strip().startswith('def '):
            # 确保sys.path在导入之前设置
            new_lines.append('import sys\n')
            new_lines.append('import os\n')
            new_lines.append('sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))\n')

            # 根据文件路径判断需要什么导入
            if 'validation' in file_path:
                new_lines.append('from modules.validation import process_excel_with_validation\n')
                if 'openpyxl' in content:
                    new_lines.append('import openpyxl\n')
            elif 'merge' in file_path:
                new_lines.append('from modules.merge import merge_excel_tables\n')

            skip_until_next = False

        new_lines.append(line)

    # 写入文件
    new_content = ''.join(new_lines)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f'✅ 已修复: {file_path}')

# 定义需要修复的文件
validation_files = [
    'tests/validation/test_standard.py',
    'tests/validation/test_abnormal_detail.py',
    'tests/validation/test_string_columns.py',
    'tests/validation/test_direct.py',
    'tests/validation/test_final.py',
    'tests/validation/verify_fix.py',
]

merge_files = [
    'tests/merge/test_basic_merge.py',
    'tests/merge/test_extra_columns.py',
    'tests/merge/test_multi_column.py',
    'tests/merge/test_merge.py',
]

for file_path in validation_files + merge_files:
    if os.path.exists(file_path):
        fix_file(file_path)

print('\n✅ 所有测试文件已修复！')
