import pandas as pd
from typing import List, Dict, Any, Optional
import os

def process_excel_with_validation(
    input_file: str,
    sheet_name: str,
    group_columns: List[str],
    compare_columns: List[str],
    output_columns: Optional[List[str]] = None,
    output_file: str = 'validation_result.xlsx',
    string_columns: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    处理Excel文件并进行数据验证

    参数:
        input_file: 输入Excel文件路径
        sheet_name: 工作表名称
        group_columns: 分组列名列表
        compare_columns: 需要比较是否相等的列名列表（必须是2列）
        output_columns: 输出到新Excel的列名列表（包含分组状态列）
        output_file: 输出文件名
        string_columns: 需要保持为字符串格式的列名列表（避免"001"变成1）

    返回:
        处理后的DataFrame
    """
    # 确保输出文件在当前文件夹
    if not os.path.isabs(output_file):
        # 如果是相对路径，确保输出到当前文件夹
        output_file = os.path.join(os.getcwd(), output_file)

    # 1. 读取Excel数据，保持原始格式
    print(f"正在读取 {input_file} 的 {sheet_name} 工作表...")
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # 如果指定了字符串列，将这些列转换为字符串以保持格式
    if string_columns:
        for col in string_columns:
            if col in df.columns:
                df[col] = df[col].astype(str)

    # 2. 数据验证检查
    print("正在验证数据...")
    # 获取比较列名
    col1, col2 = compare_columns[0], compare_columns[1]

    # 检查数据是否有效
    if len(compare_columns) != 2:
        raise ValueError("compare_columns 必须包含 exactly 2 列名")

    if col1 not in df.columns or col2 not in df.columns:
        missing_cols = [col for col in compare_columns if col not in df.columns]
        raise ValueError(f"比较列不存在: {missing_cols}")

    # 逐行比较两列数据
    df['行是否正常'] = (df[col1] == df[col2])

    # 3. 分组检查
    print("正在进行分组验证...")

    # 使用更直接的方法计算统计信息
    # 获取所有唯一的分组
    unique_groups = df.groupby(group_columns).size().reset_index()

    # 计算每个组的状态
    group_results = []
    for _, group_row in unique_groups.iterrows():
        group_key = tuple(group_row[group_columns])

        # 筛选出属于当前组的数据
        group_data = df.copy()
        for i, col in enumerate(group_columns):
            group_data = group_data[group_data[col] == group_key[i]]

        # 检查是否全正常
        is_all_normal = group_data['行是否正常'].all()
        normal_count = group_data['行是否正常'].sum()
        total_count = len(group_data)
        abnormal_count = total_count - normal_count

        # 创建结果行
        result_row = group_row.to_dict()
        result_row['验证状态'] = '正常' if is_all_normal else '异常'
        result_row['正常行数'] = normal_count
        result_row['异常行数'] = abnormal_count
        result_row['总行数'] = total_count
        result_row['异常率'] = f"{abnormal_count/total_count*100:.1f}%"

        group_results.append(result_row)

    # 创建结果DataFrame
    group_stats = pd.DataFrame(group_results)

    # 4. 处理输出列选择
    # 确保验证状态在输出列中
    if output_columns is not None:
        # 只保留分组列（因为分组后的结果不包含原始数据列）
        final_output_columns = group_columns.copy()
        # 添加验证状态列
        if '验证状态' not in final_output_columns:
            final_output_columns.append('验证状态')
    else:
        # 如果没有指定输出列，使用所有列
        final_output_columns = group_stats.columns.tolist()

    # 5. 选择最终的组级别数据（只保留汇总数据）
    final_result = group_stats[final_output_columns].copy()

    # 6. 如果用户指定了额外的输出列，尝试添加汇总信息
    # 注意：只有已经在 group_stats 中的列才能被添加
    if output_columns is not None:
        # 检查是否有在output_columns中但不在final_output_columns中的列
        additional_cols = [col for col in output_columns
                          if col in group_stats.columns and col not in final_output_columns]
        if additional_cols:
            # 添加这些汇总列
            for col in additional_cols:
                if col in group_stats.columns:
                    final_result[col] = group_stats[col]
        else:
            # 如果没有额外的列可用，显示可用的列
            print(f"\n可用的汇总列: {list(group_stats.columns)}")
            print(f"注意: 不能在汇总数据中包含原始数据列（如订单号、产品代码）")
            print(f"因为这些列在分组后不存在于汇总结果中\n")

    # 6. 处理字符串格式列
    if string_columns:
        # 确保分组列中的字符串列保持格式
        for col in group_columns:
            if col in string_columns and col in final_result.columns:
                final_result[col] = final_result[col].astype(str)

    # 7. 输出到Excel，保持原始格式
    print(f"正在输出结果到 {output_file}...")

    # 创建Excel writer，使用openpyxl保持格式
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 验证结果 - 组级别的汇总数据
        # 创建副本并强制转换为字符串格式
        final_result_copy = final_result.copy()

        if string_columns:
            for col in string_columns:
                if col in final_result_copy.columns:
                    final_result_copy[col] = final_result_copy[col].astype(str)

        final_result_copy.to_excel(writer, sheet_name='验证结果', index=False)

        # 分组统计摘要 - 更详细的统计信息
        group_summary = group_stats[group_columns + ['验证状态', '正常行数', '异常行数', '总行数', '异常率']]
        group_summary_copy = group_summary.copy()

        # 同样处理分组统计中的字符串列
        if string_columns:
            for col in string_columns:
                if col in group_summary_copy.columns:
                    group_summary_copy[col] = group_summary_copy[col].astype(str)

        group_summary_copy.to_excel(writer, sheet_name='分组统计', index=False)

        # 异常详情（如果原始数据可用）
        if '行是否正常' in df.columns:
            abnormal_data = df[~df['行是否正常']]
            if not abnormal_data.empty:
                # 使用输出列和比较列创建异常详情
                abnormal_cols = group_columns.copy()
                if output_columns:
                    for col in output_columns:
                        if col not in abnormal_cols and col != '验证状态':
                            abnormal_cols.append(col)
                # 添加比较列
                for col in compare_columns:
                    if col not in abnormal_cols:
                        abnormal_cols.append(col)
                # 添加状态列
                if '行是否正常' not in abnormal_cols:
                    abnormal_cols.append('行是否正常')

                # 确保异常详情中的字符串列格式正确
                abnormal_data_copy = abnormal_data[abnormal_cols].copy()
                if string_columns:
                    for col in string_columns:
                        if col in abnormal_data_copy.columns:
                            abnormal_data_copy[col] = abnormal_data_copy[col].astype(str)

                abnormal_data_copy.to_excel(writer, sheet_name='异常详情', index=False)

    print(f"文件已保存到: {output_file}")
    return final_result


# 使用示例
if __name__ == "__main__":
    # 创建测试数据
    test_data = {
        '订单号': ['001', '002', '003', '001', '002', '003'],
        '产品代码': ['P01', 'P02', 'P03', 'P01', 'P02', 'P03'],
        '部门': ['A', 'A', 'B', 'B', 'A', 'B'],
        '月份': ['2024-01', '2024-01', '2024-01', '2024-02', '2024-02', '2024-02'],
        '计划数量': [100, 200, 150, 120, 180, 160],
        '实际数量': [100, 200, 150, 120, 180, 160]
    }
    test_df = pd.DataFrame(test_data)
    test_df.to_excel('test_data.xlsx', index=False)

    print("测试数据创建完成")
    print(test_df)

    # 测试：不使用string_columns参数
    print("\n=== 测试1：不使用string_columns参数 ===")
    result1 = process_excel_with_validation(
        input_file='test_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门', '月份'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '月份', '订单号', '产品代码'],
        output_file='test_result_without_string.xlsx'
    )

    print("结果1:")
    print(result1)
    if '订单号' in result1.columns:
        print("订单号的数据类型:", result1['订单号'].dtype)

    # 测试：使用string_columns参数
    print("\n=== 测试2：使用string_columns参数 ===")
    result2 = process_excel_with_validation(
        input_file='test_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门', '月份'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '月份', '订单号', '产品代码'],
        output_file='test_result_with_string.xlsx',
        string_columns=['订单号', '产品代码']
    )

    print("结果2:")
    print(result2)
    if '订单号' in result2.columns:
        print("订单号的数据类型:", result2['订单号'].dtype)