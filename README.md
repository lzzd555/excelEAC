# Excel 数据验证工具

一个用于处理Excel文件并进行数据验证的Python工具，支持分组比较和异常检测。

## 功能特点

- **数据读取**：从指定Excel文件的指定工作表读取数据
- **分组处理**：按指定列进行分组
- **数据验证**：比较两列数据是否相等，标记数据正常性
- **智能判断**：根据每组中所有行是否正常，判断整组状态
- **多sheet输出**：
  - 验证结果：带验证状态的完整数据
  - 分组统计：各组的汇总统计
  - 异常详情：所有异常行的详细信息
- **灵活配置**：可自定义输出列和输出文件名

## 安装要求

```bash
pip install pandas openpyxl
```

## 使用方法

### 基本使用

```python
from excel_validator import process_excel_with_validation

result = process_excel_with_validation(
    input_file='your_data.xlsx',        # 输入Excel文件路径
    sheet_name='Sheet1',                # 工作表名称
    group_columns=['部门'],             # 分组列（可以多列）
    compare_columns=['计划值', '实际值'], # 需要比较的两列
    output_columns=['部门', '产品', '计划值', '实际值'], # 输出列（可选）
    output_file='validation_result.xlsx' # 输出文件名（可选）
)
```

### 参数说明

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| input_file | str | 是 | 输入Excel文件路径 |
| sheet_name | str | 是 | 工作表名称 |
| group_columns | List[str] | 是 | 分组列名列表（可以是一列或多列） |
| compare_columns | List[str] | 是 | 需要比较是否相等的列名列表（必须是2列） |
| output_columns | List[str] | 否 | 输出到新Excel的列名列表（默认所有列） |
| output_file | str | 否 | 输出文件名（默认validation_result.xlsx） |

### 输出文件结构

生成的Excel文件包含3个工作表：

1. **验证结果**
   - **重要**：输出的是组级别的汇总数据，行数等于组数量
   - 包含分组列和指定的输出列，以及"验证状态"列
   - 每行代表一个组的汇总信息，不是原始数据

2. **分组统计**
   - 各组的基本统计信息
   - 包括：组状态、正常行数、异常行数、总行数、异常率

3. **异常详情**
   - 所有不符合条件的数据行
   - 便于快速定位问题数据

## 使用示例

### 示例1：订单数据验证

```python
# 验证订单的计划数量和实际数量是否一致
result = process_excel_with_validation(
    input_file='orders.xlsx',
    sheet_name='订单数据',
    group_columns=['订单号'],
    compare_columns=['计划数量', '实际数量'],
    output_columns=['订单号', '产品名称', '计划数量', '实际数量', '单价']
)
```

### 示例2：库存数据验证

```python
# 验证系统库存和实际库存是否一致
result = process_excel_with_validation(
    input_file='inventory.xlsx',
    sheet_name='库存表',
    group_columns=['仓库', '分类'],
    compare_columns=['系统库存', '实际库存'],
    output_file='inventory_validation.xlsx'
)
```

### 示例3：财务数据验证

```python
# 验证预算和实际支出是否一致
result = process_excel_with_validation(
    input_file='finance.xlsx',
    sheet_name='支出记录',
    group_columns=['部门', '月份'],
    compare_columns=['预算金额', '实际支出'],
    output_columns=['部门', '月份', '项目名称', '预算金额', '实际支出']
)
```

## 注意事项

1. 确保输入Excel文件存在且可读
2. 确保比较的列名在文件中存在
3. 输出文件会自动保存在运行脚本的当前文件夹
4. compare_columns参数必须包含恰好2个列名
5. 如果某组中所有行都正常，则该组状态为"正常"，否则为"异常"

## 代码结构

```
excelEAC/
├── excel_validator.py    # 主程序文件
├── README.md            # 项目说明文档
└── .git/                # Git版本控制
```

## 许可证

此项目仅供学习和参考使用。