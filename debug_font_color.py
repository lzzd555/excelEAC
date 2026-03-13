"""
调试字体颜色复制问题
"""
import openpyxl

# 读取模板文件
wb = openpyxl.load_workbook('template_complete_styles.xlsx', data_only=False)
ws = wb.active

print("模板文件字体颜色分析:")
print("=" * 50)

# 检查字体颜色
for col_idx in range(1, 10):
    cell = ws.cell(row=2, column=col_idx)
    print(f"\n列 {col_idx}: '{cell.value}'")

    if cell.font:
        print(f"  has_font: True")
        print(f"  font.name: {cell.font.name}")
        print(f"  font.size: {cell.font.size}")
        print(f"  font.bold: {cell.font.bold}")
        print(f"  font.italic: {cell.font.italic}")

        if cell.font.color:
            print(f"  has_color: True")
            print(f"  color.rgb: {cell.font.color.rgb}")
            print(f"  color.type: {type(cell.font.color)}")
        else:
            print(f"  has_color: False")
    else:
        print(f"  has_font: False")

wb.close()