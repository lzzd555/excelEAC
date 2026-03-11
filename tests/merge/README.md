# 合并模块测试

这个目录包含表合并模块的测试代码和测试样例数据。

## 测试文件列表

- `test_merge.py` - 表合并功能测试套件
- `test_basic_merge.py` - 基本表合并测试（单列匹配）
- `test_extra_columns.py` - 带额外列的合并测试
- `test_multi_column.py` - 多列匹配测试

## 测试样例数据

### table_a_basic.xlsx
表A基础测试数据（单列匹配）
- 列：ID, 姓名, 部门, 年龄
- 数据量：4行
- 用途：测试基本合并功能

### table_b_basic.xlsx
表B基础测试数据（单列匹配）
- 列：ID, 职位, 薪资
- 数据量：4行
- 匹配关系：ID列
- 用途：与table_a_basic.xlsx配合测试

### table_a_extra.xlsx
表A额外测试数据（带额外列）
- 列：ID, 姓名, 部门, 年龄
- 数据量：3行
- 用途：测试带额外列的合并

### table_b_extra.xlsx
表B额外测试数据（带额外列）
- 列：ID, 职位, 薪资, 绩效等级
- 数据量：3行
- 用途：与table_a_extra.xlsx配合测试

### table_a_multi.xlsx
表A多列匹配测试数据
- 列：订单号, 产品代码, 客户, 数量
- 数据量：3行
- 用途：测试多列匹配

### table_b_multi.xlsx
表B多列匹配测试数据
- 列：订单号, 产品代码, 单价, 状态
- 数据量：4行
- 匹配关系：订单号, 产品代码
- 用途：与table_a_multi.xlsx配合测试

## 运行测试

### 运行合并模块测试

```bash
# 运行完整的测试套件
python tests/merge/test_merge.py
```

### 运行单个测试

```bash
# 基本合并测试
python tests/merge/test_basic_merge.py

# 带额外列的合并测试
python tests/merge/test_extra_columns.py

# 多列匹配测试
python tests/merge/test_multi_column.py
```

### 使用命令行接口

```bash
# 使用主程序进行合并
python main.py merge -a sample_data/table_a.xlsx -A Sheet1 -b sample_data/table_b.xlsx -B Sheet1 -m ID -o result.xlsx
```

## 测试覆盖的功能

### test_basic_merge.py
测试基本表合并功能（单列匹配）。
验证点：
- 单列匹配逻辑
- 基本合并功能
- 输出格式正确性
预期结果：3行匹配数据

### test_extra_columns.py
测试带额外列配置的合并。
验证点：
- table_a_extra_columns参数
- table_b_extra_columns参数
- 列顺序正确性
- 列去重功能
预期结果：7列（ID + 姓名,部门,年龄 + 职位,薪资,绩效等级）

### test_multi_column.py
测试多列匹配功能。
验证点：
- 多列匹配逻辑
- 复杂条件判断
- 结果正确性
预期结果：3行匹配数据（001-P01, 002-P02, 003-P03）

## 测试数据说明

### 单列匹配测试

使用 `table_a_basic.xlsx` 和 `table_b_basic.xlsx`：
- 匹配列：ID
- 不添加额外列
- 预期输出：3行（ID为A001, A002, A003的记录）

### 带额外列测试

使用 `table_a_extra.xlsx` 和 `table_b_extra.xlsx`：
- 匹配列：ID
- 表A额外列：姓名, 部门, 年龄
- 表B额外列：职位, 薪资, 绩效等级
- 预期输出：3行，7列数据

### 多列匹配测试

使用 `table_a_multi.xlsx` 和 `table_b_multi.xlsx`：
- 匹配列：订单号, 产品代码
- 表A额外列：客户, 数量
- 表B额外列：单价, 状态
- 预期输出：3行，6列数据

## 合并逻辑说明

### 匹配规则

1. **单列匹配**：
   - 表A的匹配列值 = 表B的匹配列值 → 匹配

2. **多列匹配**：
   - 表A的所有匹配列值 = 表B的所有匹配列值 → 匹配

### 输出列顺序

1. **匹配列**：始终在前面
2. **表A额外列**：在匹配列之后
3. **表B额外列**：在表A额外列之后
4. **自动去重**：避免重复列

### 数据处理

- 匹配时使用表A的值（避免重复）
- 额外列从对应表提取
- 保持数据类型不变
- 支持string_columns参数保持格式

## 测试输出

测试产生的临时文件会输出到 `../../test_output/` 目录。
