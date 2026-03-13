# Excel工具包

一个提供数据验证、表合并和模板生成功能的Python工具包，采用模块化设计。

## 功能特点

### 验证模块 (`modules/validation.py`)

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

### 合并模块 (`modules/merge.py`)

- **表匹配合并**：基于指定列将两张Excel表的数据合并
- **灵活配置**：
  - 指定匹配列（单列或多列）
  - 分别配置表A和表B的额外列
  - 自动去重和列排序
- **数据类型保持**：支持字符串格式保持，避免"001"变成"1"
- **多列匹配**：支持按多个列同时匹配数据
- **列名映射**：支持表A和表B使用不同列名进行匹配

### 模板生成模块 (`modules/template_generator.py`)

- **多数据源支持**：可从多个Excel文件读取数据，合并到输出文件
- **公式保留与转换**：
  - 保留模板中的公式结构
  - 自动将sheet别名转换为外部文件引用（如 `sheet0!A1` → `'[data.xlsx]SheetName'!A1`）
  - 支持复杂嵌套公式（IF、COUNTIFS、SUMIFS等）
  - 支持整列引用（A:A）和范围引用（A1:B10）
- **实际Sheet名支持**：模板中可直接使用实际sheet名（如 `'ESDP-Bpart'!A:A`），自动匹配对应文件
- **列映射功能**：支持源列名到目标列名的映射
- **两种数据写入模式**：
  - 外部引用模式（默认）：使用Excel外部引用公式，数据与源文件保持链接
  - 直接值模式（`--direct-values`）：直接写入数据值，适用于Numbers等不支持外部引用的软件
- **公式仅数据源**：支持仅在公式中引用的数据源（无需列映射）

## 安装要求

```bash
pip install pandas openpyxl
```

## 使用方法

### 方法1：直接导入模块

```python
from modules.validation import process_excel_with_validation
from modules.merge import merge_excel_tables
from modules.template_generator import generate_excel_from_template

# 使用数据验证功能
result = process_excel_with_validation(
    input_file='data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门'],
    compare_columns=['计划值', '实际值'],
    output_file='validation_result.xlsx'
)

# 使用表合并功能
# 方式1：使用相同的列名（旧格式，仍然支持）
merged = merge_excel_tables(
    table_a_file='table_a.xlsx',
    table_a_sheet='Sheet1',
    table_b_file='table_b.xlsx',
    table_b_sheet='Sheet1',
    match_columns=['ID'],
    table_a_extra_columns=['姓名', '部门'],
    table_b_extra_columns=['职位', '薪资'],
    output_file='merge_result.xlsx'
)

# 方式2：使用不同的列名映射（新功能）
merged = merge_excel_tables(
    table_a_file='table_a.xlsx',
    table_a_sheet='Sheet1',
    table_b_file='table_b.xlsx',
    table_b_sheet='Sheet1',
    match_columns={'ID': '员工编号', '部门': '部门编码'},  # 表A列名映射到表B列名
    table_a_extra_columns=['姓名', '职位'],
    table_b_extra_columns=['薪资', '入职日期'],
    output_file='merge_result.xlsx',
    string_columns=['ID', '员工编号']
)

# 使用模板生成功能
result = generate_excel_from_template(
    template_file='template.xlsx',
    template_sheet='Sheet1',
    formula_columns=['Sales', 'Cost', 'Profit'],
    data_sources=[
        {
            'file_path': 'sales.xlsx',
            'sheet_name': 'Data',
            'column_mappings': [
                {'source': 'Date', 'target': 'Date'},
                {'source': 'SalesAmt', 'target': 'Sales'}
            ],
            'alias': 'sheet0'
        },
        {
            'file_path': 'costs.xlsx',
            'sheet_name': 'Data',
            'column_mappings': [
                {'source': 'Date', 'target': 'Date'},
                {'source': 'CostAmt', 'target': 'Cost'}
            ],
            'alias': 'sheet1'
        }
    ],
    output_file='result.xlsx',
    string_columns=['Date']
)
```

### 方法2：使用主程序（命令行）

#### 数据验证命令

```bash
python main.py validate -i data.xlsx -s Sheet1 -g 部门 -c 计划值,实际值 -o result.xlsx
```

参数说明：
- `-i, --input`: 输入Excel文件路径（必需）
- `-s, --sheet`: 工作表名称（必需）
- `-g, --group-columns`: 分组列名，逗号分隔（必需）
- `-c, --compare-columns`: 比较列名，逗号分隔，2列（必需）
- `-o, --output`: 输出文件名（默认：validation_result.xlsx）
- `--output-columns`: 输出列名，逗号分隔（可选）
- `--string-columns`: 字符串列名，逗号分隔（可选）
- `--abnormal-detail-columns`: 异常详情列名，逗号分隔（可选）

#### 表合并命令

##### 使用相同的列名（旧格式）
```bash
python main.py merge -a table_a.xlsx -A Sheet1 -b table_b.xlsx -B Sheet1 -m ID -a_extra 姓名,部门 -b_extra 职位,薪资 -o merged.xlsx
```

##### 使用不同的列名映射（新格式）
```bash
python main.py merge -a table_a.xlsx -A Sheet1 -b table_b.xlsx -B Sheet1 -m "ID:员工编号,部门:部门编码" -a_extra 姓名,职位 -b_extra 薪资,入职日期 -o merged.xlsx --string-columns "ID,员工编号"
```

参数说明：
- `-a, --table-a`: 表A的Excel文件路径（必需）
- `-A, --table-a-sheet`: 表A的工作表名称（必需）
- `-b, --table-b`: 表B的Excel文件路径（必需）
- `-B, --table-b-sheet`: 表B的工作表名称（必需）
- `-m, --match-columns`: 匹配列名（必需）
  - 旧格式：`ID` 或 `ID,部门`（表A和表B使用相同列名）
  - 新格式：`"ID:员工编号,部门:部门编码"`（表A列名:表B列名）
- `--table-a-extra-columns`: 表A额外列名，逗号分隔（可选）
- `--table-b-extra-columns`: 表B额外列名，逗号分隔（可选）
- `-o, --output`: 输出文件名（默认：merge_result.xlsx）
- `--string-columns`: 字符串列名，逗号分隔（可选）

#### 模板生成命令

```bash
python main.py template -t template.xlsx -ts Sheet1 -f "Sales,Cost,Profit" \
    -d sales.xlsx Data "Date:Date,SalesAmt:Sales" sheet0 \
    -d costs.xlsx Data "Date:Date,CostAmt:Cost" sheet1 \
    -o result.xlsx
```

参数说明：
- `-t, --template`: 模板Excel文件路径（必需）
- `-ts, --template-sheet`: 模板工作表名称（必需）
- `-f, --formula-columns`: 公式列名，逗号分隔（可选）
- `-d, --data-source`: 数据源（可多次使用）。格式: `file_path sheet_name "SrcCol:TgtCol,..." alias`
- `-o, --output`: 输出文件名（默认：output.xlsx）
- `--string-columns`: 字符串列名，逗号分隔（可选）
- `--direct-values`: 直接写入数据值而非外部引用公式（适用于Numbers等不支持外部引用的软件）

## 验证模块使用示例

### 示例1：订单数据验证

```python
from modules.validation import process_excel_with_validation

result = process_excel_with_validation(
    input_file='orders.xlsx',
    sheet_name='订单数据',
    group_columns=['订单号'],
    compare_columns=['计划数量', '实际数量'],
    output_columns=['订单号', '验证状态', '总行数'],
    output_file='validation_result.xlsx',
    string_columns=['订单号']
)
```

### 示例2：带异常详情配置的验证

```python
from modules.validation import process_excel_with_validation

result = process_excel_with_validation(
    input_file='data.xlsx',
    sheet_name='Sheet1',
    group_columns=['部门'],
    compare_columns=['计划金额', '实际金额'],
    abnormal_detail_columns=['订单号', '产品代码', '计划金额', '实际金额'],  # 指定异常详情列
    output_file='result.xlsx',
    string_columns=['订单号', '产品代码']
)
```

## 合并模块使用示例

### 示例1：基本表合并（单列匹配）

```python
from modules.merge import merge_excel_tables

result = merge_excel_tables(
    table_a_file='employees.xlsx',
    table_a_sheet='员工信息',
    table_b_file='salaries.xlsx',
    table_b_sheet='薪资信息',
    match_columns=['员工ID'],  # 按员工ID匹配
    table_a_extra_columns=['姓名', '部门'],  # 从表A添加
    table_b_extra_columns=['职位', '薪资'],  # 从表B添加
    output_file='merged_employees.xlsx'
)

# 结果：匹配的员工将包含姓名、部门、职位、薪资信息
```

### 示例2：表合并（多列匹配）

```python
from modules.merge import merge_excel_tables

result = merge_excel_tables(
    table_a_file='orders.xlsx',
    table_a_sheet='订单信息',
    table_b_file='products.xlsx',
    table_b_sheet='产品信息',
    match_columns=['订单号', '产品代码'],  # 按订单号和产品代码同时匹配
    table_a_extra_columns=['客户', '数量'],  # 从表A添加
    table_b_extra_columns=['单价', '状态'],  # 从表B添加
    output_file='merged_orders.xlsx',
    string_columns=['订单号', '产品代码']
)

# 结果：同时匹配订单号和产品代码的行将被合并
```

### 示例3：只合并匹配列

```python
from modules.merge import merge_excel_tables

result = merge_excel_tables(
    table_a_file='table_a.xlsx',
    table_a_sheet='Sheet1',
    table_b_file='table_b.xlsx',
    table_b_sheet='Sheet1',
    match_columns=['ID', '日期'],  # 按ID和日期匹配
    # 不添加额外列
    table_a_extra_columns=None,
    table_b_extra_columns=None,
    output_file='simple_merge.xlsx'
)

# 结果：只包含匹配列（ID, 日期）
```

## 模板生成模块使用示例

### 示例1：基本模板生成（使用别名）

```python
from modules.template_generator import generate_excel_from_template

# 模板中的公式: =sheet0!B2+sheet1!B2
result = generate_excel_from_template(
    template_file='template.xlsx',
    template_sheet='Sheet1',
    formula_columns=['Total'],  # 包含公式的列
    data_sources=[
        {
            'file_path': 'sales.xlsx',
            'sheet_name': '销售数据',
            'column_mappings': [
                {'source': 'SalesAmt', 'target': 'Sales'}
            ],
            'alias': 'sheet0'  # 模板中引用的别名
        },
        {
            'file_path': 'costs.xlsx',
            'sheet_name': '成本数据',
            'column_mappings': [
                {'source': 'CostAmt', 'target': 'Cost'}
            ],
            'alias': 'sheet1'
        }
    ],
    output_file='result.xlsx'
)

# 输出公式: ='[sales.xlsx]销售数据'!B2+'[costs.xlsx]成本数据'!B2
```

### 示例2：使用实际Sheet名（无需别名）

```python
# 模板中的公式: ='ESDP-Bpart'!A:A
result = generate_excel_from_template(
    template_file='template.xlsx',
    template_sheet='分析',
    formula_columns=['计算结果'],
    data_sources=[
        {
            'file_path': 'data_a.xlsx',
            'sheet_name': 'ESDP-Bpart',  # 直接使用实际sheet名
            'column_mappings': []  # 可以为空，仅用于公式引用
        }
    ],
    output_file='result.xlsx'
)

# 输出公式: ='[data_a.xlsx]ESDP-Bpart'!A:A
```

### 示例3：复杂嵌套公式

```python
# 模板中的复杂公式:
# =IF(B2="新发货",
#     SUMIFS(sheet0!C:C,sheet0!A:A,C2,sheet0!B:B,D2),
#     SUMIFS(sheet1!C:C,sheet1!A:A,C2,sheet1!B:B,D2))

result = generate_excel_from_template(
    template_file='template.xlsx',
    template_sheet='加载状态分析',
    formula_columns=['计算结果'],
    data_sources=[
        {
            'file_path': 'bpart.xlsx',
            'sheet_name': 'ESDP-Bpart',
            'column_mappings': [],
            'alias': 'sheet0'
        },
        {
            'file_path': 'cpart.xlsx',
            'sheet_name': 'ESDP-Cpart',
            'column_mappings': [],
            'alias': 'sheet1'
        }
    ],
    output_file='result.xlsx'
)

# 公式会被正确转换为外部文件引用
```

### 示例4：直接值模式（Numbers兼容）

```bash
# 使用 --direct-values 参数直接写入数据值
python main.py template -t template.xlsx -ts Sheet1 \
    -d sales.xlsx Data "SalesAmt:Sales" sheet0 \
    -o result.xlsx --direct-values
```

```

## 参数说明

### 验证模块参数

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

### 合并模块参数

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| table_a_file | str | 是 | 表A的Excel文件路径 |
| table_a_sheet | str | 是 | 表A的工作表名称 |
| table_b_file | str | 是 | 表B的Excel文件路径 |
| table_b_sheet | str | 是 | 表B的工作表名称 |
| match_columns | List[str] | 是 | 需要匹配的列名列表（可以是单列或多列） |
| table_a_extra_columns | List[str] | 否 | 从表A中额外添加的列名列表（除匹配列外） |
| table_b_extra_columns | List[str] | 否 | 从表B中额外添加的列名列表（除匹配列外） |
| output_file | str | 否 | 输出文件名（默认merge_result.xlsx） |
| string_columns | List[str] | 否 | 需要保持为字符串格式的列名列表（避免"001"变成1） |

### 模板生成模块参数

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| template_file | str | 是 | 模板Excel文件路径 |
| template_sheet | str | 是 | 模板工作表名称 |
| formula_columns | List[str] | 否 | 包含公式的列名列表 |
| data_sources | List[Dict] | 是 | 数据源配置列表 |
| output_file | str | 否 | 输出文件名（默认output.xlsx） |
| string_columns | List[str] | 否 | 需要保持为字符串格式的列名列表 |
| use_external_refs | bool | 否 | 是否使用外部引用（默认True），False则直接写入数据值 |

**数据源配置（data_sources中的每个元素）：**

| 字段 | 类型 | 必需 | 描述 |
|------|------|------|------|
| file_path | str | 是 | 数据源Excel文件路径 |
| sheet_name | str | 是 | 数据源工作表名称 |
| column_mappings | List[Dict] | 否 | 列映射配置（可以为空，仅用于公式引用） |
| alias | str | 否 | 别名（如"sheet0"，用于模板公式中引用） |

**列映射配置（column_mappings中的每个元素）：**

| 字段 | 类型 | 描述 |
|------|------|------|
| source | str | 源列名（数据源中的列名） |
| target | str | 目标列名（模板中的列名） |

## 合并逻辑说明

1. **匹配列处理**：
   - 根据指定的匹配列（可以是一个或多个）在表A和表B中查找匹配的行
   - 所有匹配列的值都必须相同才认为匹配

2. **额外列处理**：
   - 表A额外列：从表A的匹配行中提取，添加到合并结果
   - 表B额外列：从表B的匹配行中提取，添加到合并结果
   - 如果表A和表B都有同名列（非匹配列），优先使用表A的值

3. **列顺序**：
   - 输出列顺序：匹配列 + 表A额外列 + 表B额外列
   - 自动去重，避免重复列

## 输出文件结构

### 验证模块输出

生成的Excel文件包含3个工作表：

1. **验证结果**
   - 分组级别的汇总数据
   - 包含分组列和指定的输出列，以及"验证状态"列

2. **分组统计**
   - 各组的基本统计信息
   - 包括：组状态、正常行数、异常行数、总行数、异常率

3. **异常详情**
   - 包含所有行（正常和异常）
   - 添加"是否异常"列标识异常状态
   - 异常行有红色背景色标记

### 合并模块输出

生成的Excel文件包含1个工作表：

1. **合并结果**
   - 按匹配列合并的数据
   - 包含：匹配列 + 表A额外列 + 表B额外列
   - 自动去重

## 数据格式保持

### 问题：数据格式丢失

当处理包含编号、代码等数据时，pandas可能会改变数据格式：
- `001` → `1` （前导零丢失）
- `00123` → `123` （前导零丢失）

### 解决方案：使用 `string_columns` 参数

```python
# 保持订单号、产品代码等数据的原始格式
result = merge_excel_tables(
    table_a_file='orders.xlsx',
    table_a_sheet='订单数据',
    table_b_file='products.xlsx',
    table_b_sheet='产品数据',
    match_columns=['订单号'],
    string_columns=['订单号', '产品代码']  # 这些列将保持字符串格式
)
```

## 代码结构

```
excelEAC/
├── main.py                   # 主程序入口
├── modules/                  # 功能模块
│   ├── __init__.py          # 模块包初始化
│   ├── validation.py         # 数据验证模块
│   ├── merge.py             # 表合并模块
│   └── template_generator.py # 模板生成模块
├── tests/                   # 测试代码
│   ├── README.md            # 测试说明
│   ├── validation/          # 验证模块测试
│   │   ├── README.md        # 验证测试说明
│   │   └── test_*.py        # 验证测试文件
│   ├── merge/               # 合并模块测试
│   │   ├── README.md        # 合并测试说明
│   │   ├── sample_data/     # 测试样例数据
│   │   └── test_*.py        # 合并测试文件
│   └── template/            # 模板生成模块测试
│       ├── test_template_generator.py  # 基础功能测试
│       ├── test_complex_formula.py     # 复杂公式测试
│       └── test_real_sheet_name.py     # 实际sheet名测试
├── run_tests.py             # 测试运行脚本
├── README.md                # 项目说明文档
└── .gitignore               # Git忽略配置
```

## 运行测试

```bash
# 运行验证模块测试
python tests/test_standard.py

# 运行合并模块测试
python tests/test_merge.py

# 运行模板生成模块测试
python tests/template/test_template_generator.py
python tests/template/test_complex_formula.py
python tests/template/test_real_sheet_name.py

# 运行所有测试
python run_tests.py
```

## 注意事项

1. 确保输入Excel文件存在且可读
2. 验证模块：确保比较的列名在文件中存在
3. 合并模块：确保匹配列在两个表中都存在
4. 合并模块：确保额外列在对应的表中存在
5. **重要**：对于需要保持格式（如"001"）的列，请使用 `string_columns` 参数指定
6. 验证模块输出的是组级别的汇总数据，行数等于组数量，不是原始数据行数
7. **模板生成模块**：
   - 公式中的sheet名可以使用别名（sheet0, sheet1）或实际sheet名
   - 支持复杂嵌套公式和整列引用（A:A）
   - 使用 `--direct-values` 参数可直接写入数据值（适用于Numbers等不支持外部引用的软件）
   - 数据源可以只用于公式引用（column_mappings为空）

## 许可证

此项目仅供学习和参考使用。
