"""
Excel工具模块包
包含数据验证、表合并和模板生成功能
"""

from .validation import process_excel_with_validation
from .merge import merge_excel_tables
from .template_generator import generate_excel_from_template

__all__ = ['process_excel_with_validation', 'merge_excel_tables', 'generate_excel_from_template']

__version__ = '1.1.0'
