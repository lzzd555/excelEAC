#!/usr/bin/env python3
"""批量修复所有测试文件的import路径问题"""

import os

# 定义需要修复的文件
files_to_fix = {
    'tests/validation/test_standard.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
import openpyxl
''',
    'tests/validation/test_abnormal_detail.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
''',
    'tests/validation/test_string_columns.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
''',
    'tests/validation/test_direct.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
''',
    'tests/validation/test_final.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
''',
    'tests/validation/verify_fix.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.validation import process_excel_with_validation
''',
    'tests/merge/test_basic_merge.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables
''',
    'tests/merge/test_extra_columns.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables
''',
    'tests/merge/test_multi_column.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables
''',
    'tests/merge/test_merge.py': '''import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.merge import merge_excel_tables
''',
}

for file_path, new_imports in files_to_fix.items():
    if os.path.exists(file_path):
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_imports)
        print(f'✅ 已修复: {file_path}')
    else:
        print(f'⏭️ 文件不存在: {file_path}')

print('\n✅ 所有测试文件已修复！')
