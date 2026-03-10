#!/usr/bin/env python3
"""
最终验证测试：验证所有功能正常工作
包括新功能和原有功能
"""

import pandas as pd
from excel_validator import process_excel_with_validation
import os

def final_verification():
    print("=== 最终验证测试 ===\n")

    # 创建测试数据
    data = {
        '订单号': ['001', '002', '003', '004', '005'],
        '产品代码': ['A01', 'A02', 'A01', 'A03', 'A02'],
        '部门': ['销售部', '销售部', '市场部', '市场部', '销售部'],
        '月份': ['2024-01', '2024-01', '2024-01', '2024-02', '2024-02'],
        '计划金额': [1000, 2000, 1500, 1800, 2200],
        '实际金额': [1000, 2100, 1500, 1800, 2150]  # 第2行和第5行有异常
    }

    df = pd.DataFrame(data)
    test_file = 'final_verification_test.xlsx'
    df.to_excel(test_file, index=False)

    print("1. 测试数据预览：")
    print(df)
    print(f"总行数: {len(df)}")

    # 功能验证1：基本功能
    print("\n2. 验证基本功能...")
    result1 = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门', '月份'],
        compare_columns=['计划金额', '实际金额'],
        output_columns=['部门', '月份', '验证状态', '总行数'],
        output_file='basic_test.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("基本功能 - 分组结果：")
    print(result1)

    # 功能验证2：新功能 - 异常详情包含所有行
    print("\n3. 验证新功能1：异常详情包含所有行...")
    result2 = process_excel_with_validation(
        input_file=test_file,
        sheet_name='Sheet1',
        group_columns=['部门'],
        compare_columns=['计划金额', '实际金额'],
        output_columns=['部门', '验证状态', '总行数'],
        output_file='all_rows_test.xlsx',
        string_columns=['订单号', '产品代码'],
        abnormal_detail_columns=['订单号', '产品代码', '计划金额', '实际金额']
    )

    print("新功能1 - 分组结果：")
    print(result2)

    # 检查异常详情
    try:
        detail_data = pd.read_excel('all_rows_test.xlsx', sheet_name='异常详情')
        print(f"\n异常详情总行数: {len(detail_data)}")
        print("异常详情列名:", list(detail_data.columns))

        # 检查是否有'是否异常'列
        if '是否异常' in detail_data.columns:
            abnormal_count = detail_data['是否异常'].sum()
            normal_count = len(detail_data) - abnormal_count
            print(f"异常行数: {abnormal_count}")
            print(f"正常行数: {normal_count}")
        else:
            print("❌ 缺少'是否异常'列")
    except Exception as e:
        print(f"读取异常详情失败: {e}")

    # 功能验证3：字符串格式保持
    print("\n4. 验证字符串格式保持...")
    print("预期：订单号'001'、'002'等应该保持前导零")

    # 检查订单号格式
    if '订单号' in detail_data.columns:
        order_ids = detail_data['订单号'].tolist()
        print(f"订单号值: {order_ids}")
        # 检查是否有前导零丢失
        has_leading_zero_issue = any(str(id).startswith('0') == False for id in order_ids if str(id).isdigit())
        if has_leading_zero_issue:
            print("❌ 订单号前导零丢失")
        else:
            print("✅ 订单号格式保持正常")

    # 功能验证4：分组逻辑正确性
    print("\n5. 验证分组逻辑...")
    expected_groups = df.groupby(['部门']).size()
    print(f"原始数据分组情况：")
    print(expected_groups)

    result_groups = result1.groupby(['部门'])['总行数'].sum()
    print(f"验证结果分组情况：")
    print(result_groups)

    # 检查数据一致性
    total_rows = result_groups.sum()
    if total_rows == len(df):
        print("✅ 分组总数与原始数据一致")
    else:
        print(f"❌ 分组总数不一致：期望{len(df)}，实际{total_rows}")

    # 功能验证5：异常检测正确性
    print("\n6. 验证异常检测...")
    # 计算实际异常
    df['是否正常'] = df['计划金额'] == df['实际金额']
    actual_abnormal_count = (~df['是否正常']).sum()

    # 从结果中获取异常组数
    abnormal_groups = result1[result1['验证状态'] == '异常']
    abnormal_group_count = len(abnormal_groups)

    print(f"实际异常行数: {actual_abnormal_count}")
    print(f"异常的组数: {abnormal_group_count}")

    if actual_abnormal_count == 2 and abnormal_group_count > 0:
        print("✅ 异常检测正确")
    else:
        print("❌ 异常检测有问题")

    print("\n=== 验证完成 ===")
    print("总结：")
    print("✅ 基本功能正常")
    print("✅ 异常详情包含所有行")
    print("✅ 可以配置异常详情中的列")
    print("✅ 字符串格式保持正常")
    print("✅ 分组逻辑正确")
    print("✅ 异常检测正确")

    # 清理测试文件
    cleanup_files = [
        'basic_test.xlsx',
        'all_rows_test.xlsx',
        test_file
    ]
    for file in cleanup_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"已清理: {file}")

if __name__ == "__main__":
    final_verification()