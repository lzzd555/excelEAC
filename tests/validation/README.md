# 验证模块测试

这个目录包含数据验证模块的测试代码。

## 测试文件列表

- `test_abnormal_detail.py` - 用户的标准测试用例，测试异常详情中的字符串格式保持
- `test_standard.py` - 标准测试用例，基于test_abnormal_detail.py
- `test_direct.py` - 直接测试string_columns效果
- `test_final.py` - 最终验证测试，验证string_columns的有效范围
- `test_realistic.py` - 现实场景测试
- `test_string_columns.py` - string_columns参数测试
- `test_report.md` - 测试报告
- `debug_test.py` - 调试测试
- `verify_fix.py` - 验证修复

## 运行测试

### 单个测试

```bash
# 运行标准测试
python tests/validation/test_standard.py

# 运行用户标准测试
python tests/validation/test_abnormal_detail.py

# 运行string_columns测试
python tests/validation/test_string_columns.py
```

### 运行所有验证测试

```bash
# 从项目根目录运行
python main.py validate --help

# 或直接运行单个测试文件
python tests/validation/test_standard.py
```

## 测试覆盖的功能

- ✅ 字符串格式保持（前导零）
- ✅ 异常详情包含所有行
- ✅ 可配置异常详情中的列
- ✅ 分组逻辑正确性
- ✅ 异常检测正确性
- ✅ 颜色标记功能
- ✅ output_columns参数功能
- ✅ abnormal_detail_columns参数功能

## 测试说明

### test_standard.py
标准测试用例，基于用户提供的测试数据。
验证点：
- 数据验证功能
- 分组逻辑
- 异常详情包含所有行
- 字符串格式保持

### test_abnormal_detail.py
用户指定的标准测试用例。
验证点：
- 异常详情中的字符串格式
- 订单号和产品代码保持前导零

### test_merge_with_extra_columns.py
测试带额外列配置的验证。
验证点：
- abnormal_detail_columns参数
- 自定义异常详情列

### test_direct.py
直接测试string_columns效果。
验证点：
- string_columns参数是否正确工作
- 前导零是否保持

### test_final.py
最终验证测试，验证string_columns的有效范围。
验证点：
- 各种数据类型的string_columns处理
- 边界情况测试

### test_realistic.py
现实场景测试，模拟真实使用情况。
验证点：
- 大数据量处理
- 复杂分组场景
- 多列比较

### test_string_columns.py
专门测试string_columns参数功能。
验证点：
- 单列字符串保持
- 多列字符串保持
- 混合类型列处理

### debug_test.py
调试测试，用于定位问题。
验证点：
- 数据流追踪
- 中间结果检查

### verify_fix.py
验证修复效果的测试。
验证点：
- 确认问题已修复
- 回归测试

## 测试数据

测试数据临时文件输出到 `../../test_output/` 目录。
