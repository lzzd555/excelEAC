# 测试代码目录

这个目录包含所有测试代码和测试文件。

## 验证模块测试

- `test_abnormal_detail.py` - 用户的标准测试用例，测试异常详情中的字符串格式保持
- `test_standard.py` - 标准测试用例，基于test_abnormal_detail.py
- `test_direct.py` - 直接测试string_columns效果
- `test_final.py` - 最终验证测试，验证string_columns的有效范围
- `test_realistic.py` - 现实场景测试
- `test_string_columns.py` - string_columns参数测试
- `test_report.md` - 测试报告

## 合并模块测试

- `test_merge.py` - 表合并功能测试套件

## 运行测试

### 验证模块测试

```bash
# 运行标准测试
python tests/test_standard.py

# 运行用户标准测试
python tests/test_abnormal_detail.py

# 运行string_columns测试
python tests/test_string_columns.py
```

### 合并模块测试

```bash
# 运行表合并测试
python tests/test_merge.py
```

### 运行所有测试

```bash
# 使用主程序运行
python main.py validate --help
python main.py merge --help
```

## 测试输出

测试产生的临时文件会输出到 `../test_output/` 目录。

## 测试覆盖的功能

### 验证模块
- ✅ 字符串格式保持（前导零）
- ✅ 异常详情包含所有行
- ✅ 可配置异常详情中的列
- ✅ 分组逻辑正确性
- ✅ 异常检测正确性
- ✅ 颜色标记功能
- ✅ output_columns参数功能

### 合并模块
- ✅ 基本表合并
- ✅ 单列匹配
- ✅ 多列匹配
- ✅ 额外列配置
- ✅ 字符串格式保持
