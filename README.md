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
  - 异常详情：包含所有行（正常和异常），并用颜色标记异常行
- **灵活配置**：可自定义输出列、输出文件名、异常详情中的列
- **字符串格式保持**：保持前导零和特殊字符串格式
- **异常行标记**：异常行有红色背景色标记

## 安装要求

```bash
pip install pandas openpyxl
```

## 新功能说明

### 1. 异常详情包含所有行
异常详情sheet现在包含所有行（正常和异常），并通过添加"是否异常"列来标记每行的状态：
- `是`（或True）：该行数据异常
- `否`（或False）：该行数据正常

异常行会有红色背景标记，方便快速识别。

### 2. 可配置异常详情中的列
通过`abnormal_detail_columns`参数，可以指定异常详情中要显示哪些原表列：

```python
# 指定异常详情中只显示特定列
result = process_excel_with_validation(
    input_file='data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门'],
    compare_columns=['计划金额', '实际金额'],
    abnormal_detail_columns=['订单号', '产品代码', '计划金额', '实际金额'],  # 只显示这些列
    string_columns=['订单号', '产品代码']
)
```

如果不指定`abnormal_detail_columns`，系统会自动包含：
- 所有分组列
- 所有字符串列（通过`string_columns`指定）
- 所有比较列
- 状态列（行是否正常、是否异常）

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

### 使用新功能示例

```python
# 示例1：使用新功能 - 异常详情包含所有行，并配置显示特定列
from excel_validator import process_excel_with_validation

# 创建测试数据
data = {
    '订单号': ['001', '002', '003', '004'],
    '产品代码': ['A01', 'A02', 'A01', 'A03'],
    '部门': ['销售部', '销售部', '市场部', '市场部'],
    '计划金额': [1000, 2000, 1500, 1800],
    '实际金额': [1000, 2100, 1500, 1800]  # 第2行有异常
}

df = pd.DataFrame(data)
df.to_excel('test_data.xlsx', index=False)

# 运行验证，并指定异常详情中要显示的列
result = process_excel_with_validation(
    input_file='test_data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门'],
    compare_columns=['计划金额', '实际金额'],
    output_columns=['部门', '验证状态', '总行数'],
    output_file='enhanced_result.xlsx',
    string_columns=['订单号', '产品代码'],
    abnormal_detail_columns=['订单号', '产品代码', '计划金额', '实际金额']  # 只显示这些列
)

# 输出结果：
# 1. 验证结果：部门级别的汇总
# 2. 分组统计：详细的分组统计
# 3. 异常详情：包含所有行（正常和异常），标记异常行，只显示指定的列
```
```

### 参数说明

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| input_file | str | 是 | 输入Excel文件路径 |
| sheet_name | str | 是 | 工作表名称 |
| group_columns | List[str] | 是 | 分组列名列表（可以是一列或多列） |
| compare_columns | List[str] | 是 | 需要比较是否相等的列名列表（必须是2列） |
| output_columns | List[str] | 否 | 输出到新Excel的列名列表 |
| output_file | str | 否 | 输出文件名（默认validation_result.xlsx） |
| string_columns | List[str] | 否 | 需要保持为字符串格式的列名列表（避免"001"变成1） |
| abnormal_detail_columns | List[str] | 否 | 异常详情中需要显示的原表列名列表。如果为None，则自动包含分组列、比较列和字符串列。 |

### output_columns 参数详解

`output_columns` 参数控制最终输出中显示哪些列：

#### 可用的列：
1. **分组列**（必须）：如 `['部门', '月份']`
2. **汇总列**（可选）：
   - `验证状态`：'正常' 或 '异常'
   - `正常行数`：该组中正常数据的行数
   - `异常行数`：该组中异常数据的行数
   - `总行数`：该组的总行数
   - `异常率`：异常行数的百分比

#### 使用示例：
```python
# 只显示分组列和状态
output_columns=['部门', '月份']

# 显示分组列和详细统计
output_columns=['部门', '月份', '验证状态', '总行数', '正常行数']

# 显示所有可用的汇总信息
output_columns=['部门', '月份', '验证状态', '总行数', '正常行数', '异常行数', '异常率']
```

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

## 数据格式保持

### 问题：数据格式丢失

当处理包含编号、代码等数据时，pandas可能会改变数据格式：
- `001` → `1` （前导零丢失）
- `00123` → `123` （前导零丢失）

### 解决方案：使用 `string_columns` 参数

```python
# 保持订单号、产品代码等数据的原始格式
result = process_excel_with_validation(
    input_file='orders.xlsx',
    sheet_name='订单数据',
    group_columns=['客户ID'],
    compare_columns=['计划金额', '实际金额'],
    output_columns=['客户ID', '订单号', '产品代码', '计划金额', '实际金额'],
    string_columns=['订单号', '产品代码']  # 这些列将保持字符串格式
)
```

### 适用场景

- 订单编号：`ORD001` 保持为 `ORD001`
- 产品代码：`P00123` 保持为 `P00123`
- 工单编号：`WO2023001` 保持为 `WO2023001`
- 账户编号：`ACC001` 保持为 `ACC001`
- 任何需要保持前导零或固定格式的数据

### 技术实现

- 使用 `string_columns` 参数指定需要保持字符串格式的列
- 读取时将这些列转换为字符串
- 写入时保持字符串格式，使用openpyxl引擎
- 避免pandas的自动类型转换

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

### 示例4：自定义输出列（控制显示哪些汇总信息）

```python
# 只显示分组列和验证状态
result = process_excel_with_validation(
    input_file='orders.xlsx',
    sheet_name='订单数据',
    group_columns=['客户ID'],
    compare_columns=['计划金额', '实际金额'],
    output_columns=['客户ID', '验证状态']
)

# 显示分组列和详细统计
result = process_excel_with_validation(
    input_file='orders.xlsx',
    sheet_name='订单数据',
    group_columns=['客户ID'],
    compare_columns=['计划金额', '实际金额'],
    output_columns=['客户ID', '验证状态', '总行数', '正常行数', '异常率']
)

# 同时保持数据格式
result = process_excel_with_validation(
    input_file='orders.xlsx',
    sheet_name='订单数据',
    group_columns=['客户ID'],
    compare_columns=['计划金额', '实际金额'],
    output_columns=['客户ID', '验证状态', '总行数'],
    string_columns=['客户ID']  # 保持客户ID的原始格式
)
# 结果：客户ID"001"会保持为"001"而不是变成1
```

### 重要说明：为什么订单号/产品代码不出现？

**核心概念**：这个工具输出的是**分组级别的汇总数据**，不是原始数据。

#### 示例：
```
原始数据：
订单号  部门  计划  实际
001    A    100  100
002    A    200  200
003    B    150  150

分组后（输出结果）：
部门  验证状态  总行数
A    正常     2
B    正常     1
```

- **不能**在输出结果中看到订单号，因为一个组包含多个订单号
- **只能**看到：分组列 + 统计信息
- 如果需要查看原始订单数据，请查看 `异常详情` sheet

## 注意事项

1. 确保输入Excel文件存在且可读
2. 确保比较的列名在文件中存在
3. 输出文件会自动保存在运行脚本的当前文件夹
4. compare_columns参数必须包含恰好2个列名
5. 如果某组中所有行都正常，则该组状态为"正常"，否则为"异常"
6. **重要**：对于需要保持格式（如"001"）的列，请使用 `string_columns` 参数指定
7. 输出的是组级别的汇总数据，行数等于组数量，不是原始数据行数

## 代码结构

```
excelEAC/
├── excel_validator.py    # 主程序文件
├── README.md            # 项目说明文档
└── .git/                # Git版本控制
```

## 许可证

此项目仅供学习和参考使用。