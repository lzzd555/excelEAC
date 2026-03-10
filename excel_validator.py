import pandas as pd
from typing import List, Dict, Any, Optional
import os

def process_excel_with_validation(
    input_file: str,
    sheet_name: str,
    group_columns: List[str],
    compare_columns: List[str],
    output_columns: Optional[List[str]] = None,
    output_file: str = 'validation_result.xlsx'
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

    返回:
        处理后的DataFrame
    """
    # 确保输出文件在当前文件夹
    if not os.path.isabs(output_file):
        # 如果是相对路径，确保输出到当前文件夹
        output_file = os.path.join(os.getcwd(), output_file)

    # 1. 读取Excel数据
    print(f"正在读取 {input_file} 的 {sheet_name} 工作表...")
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # 如果没有指定输出列，默认使用所有列
    if output_columns is None:
        output_columns = df.columns.tolist()
    else:
        # 确保分组列在输出列中
        for col in group_columns:
            if col not in output_columns:
                output_columns.append(col)

        # 添加状态列
        status_col = '验证状态'
        if status_col not in output_columns:
            output_columns.append(status_col)

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
    # 分组检查每组是否全正常
    def check_group(group):
        is_all_normal = group['行是否正常'].all()
        return pd.Series({
            '组状态': '正常' if is_all_normal else '异常',
            '正常行数': group['行是否正常'].sum(),
            '异常行数': (~group['行是否正常']).sum(),
            '总行数': len(group),
            '异常率': f"{(~group['行是否正常']).mean()*100:.1f}%"
        })

    # 分组计算
    group_stats = df.groupby(group_columns).apply(check_group).reset_index()

    # 4. 处理输出列选择
    # 确保验证状态在输出列中
    if output_columns is not None:
        # 确保分组列在输出列中
        final_output_columns = group_columns.copy()
        # 添加其他指定的输出列
        for col in output_columns:
            if col not in final_output_columns and col != '验证状态':
                final_output_columns.append(col)
        # 添加验证状态列
        if '验证状态' not in final_output_columns:
            final_output_columns.append('验证状态')
    else:
        # 如果没有指定输出列，使用所有列
        final_output_columns = group_stats.columns.tolist()

    # 5. 重命名列名以匹配输出需求
    # 将'组状态'重命名为'验证状态'
    group_stats = group_stats.rename(columns={'组状态': '验证状态'})

    # 6. 选择最终的组级别数据（只保留汇总数据）
    final_result = group_stats[final_output_columns].copy()

    # 7. 输出到Excel
    print(f"正在输出结果到 {output_file}...")

    # 创建Excel writer，支持多个sheet
    with pd.ExcelWriter(output_file) as writer:
        # 验证结果 - 组级别的汇总数据
        final_result.to_excel(writer, sheet_name='验证结果', index=False)

        # 分组统计摘要 - 更详细的统计信息
        group_summary = group_stats[group_columns + ['验证状态', '正常行数', '异常行数', '总行数', '异常率']]
        group_summary.to_excel(writer, sheet_name='分组统计', index=False)

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

                abnormal_data[abnormal_cols].to_excel(writer, sheet_name='异常详情', index=False)

    print(f"文件已保存到: {output_file}")
    return final_result


# 使用示例
if __name__ == "__main__":
    # 示例1：基本用法
    result = process_excel_with_validation(
        input_file='your_data.xlsx',
        sheet_name='Sheet1',
        group_columns=['部门', '月份'],
        compare_columns=['计划数量', '实际数量'],
        output_columns=['部门', '月份', '产品名称'],  # 注意：输出列数会等于组数量，不是原始行数
        output_file='validation_result.xlsx'
    )

    # 示例2：自动在当前文件夹创建文件
    result2 = process_excel_with_validation(
        input_file='sales_data.xlsx',
        sheet_name='订单数据',
        group_columns=['客户ID'],
        compare_columns=['订单金额', '实付金额'],
        output_file='department_validation.xlsx'  # 会自动保存到当前文件夹
    )

    # 示例3：查看当前文件夹的输出文件
    print("\n当前文件夹中的输出文件:")
    for file in os.listdir('.'):
        if file.endswith('.xlsx'):
            print(f"- {file}")

    # 查看输出结果（验证结果sheet只有组数量行）
    print(f"\n验证结果包含 {len(result)} 行数据（等于组数量）")


# 使用说明
"""
使用说明：

1. 基本使用方法：
   from excel_validator import process_excel_with_validation

   result = process_excel_with_validation(
       input_file='your_data.xlsx',        # 你的Excel文件
       sheet_name='Sheet1',                # 工作表名称
       group_columns=['部门'],             # 分组列（可以多列）
       compare_columns=['计划值', '实际值'], # 需要比较的两列
       output_columns=['部门', '产品', '计划值', '实际值'], # 输出列（可选）
       output_file='result.xlsx'           # 输出文件名（可选）
   )

2. 参数说明：
   - input_file: 输入Excel文件路径
   - sheet_name: 工作表名称
   - group_columns: 分组列名列表（可以是一列或多列）
   - compare_columns: 需要比较是否相等的列名列表（必须是2列）
   - output_columns: 输出到新Excel的列名列表（可选，默认所有列）
   - output_file: 输出文件名（可选，默认validation_result.xlsx）

3. 输出文件包含3个sheet：
   - 验证结果：带验证状态的完整数据
   - 分组统计：各组的汇总统计
   - 异常详情：所有异常行的详细信息

4. 注意事项：
   - 确保输入文件存在
   - 确保比较的列名在文件中存在
   - 输出文件会自动保存在当前文件夹
   - 需要安装pandas和openpyxl库
"""