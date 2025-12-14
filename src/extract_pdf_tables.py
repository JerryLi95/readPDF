"""
数据提取工具
功能：
1. 提取PDF中样品数据
2. 提取CSV中样本数据
"""

import os
import sys
import pdfplumber
import pandas as pd
from pathlib import Path


def find_pdf_files(directory):
    """
    查找指定目录下的所有PDF文件
    
    Args:
        directory: 要搜索的目录路径
        
    Returns:
        PDF文件路径列表
    """
    pdf_files = []
    input_path = Path(directory)
    
    # 检查目录是否存在
    if not input_path.exists():
        print(f"错误: 路径不存在: {directory}")
        return []
    
    if not input_path.is_dir():
        print(f"错误: 不是有效的目录路径: {directory}")
        return []
    
    # 查找PDF文件
    for file in input_path.glob('*.pdf'):
        pdf_files.append(str(file))
    
    return sorted(pdf_files)


def extract_table_from_pdf(pdf_path):
    """
    从PDF文件中提取表格数据
    
    Args:
        pdf_path: PDF文件路径
        
    Returns:
        提取的表格数据（DataFrame），如果未找到则返回None
    """
    print(f"\n正在处理文件: {os.path.basename(pdf_path)}")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 遍历所有页面查找表格
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"  检查页面 {page_num}...")
                
                # 方法1: 尝试使用更宽松的表格提取设置
                table_settings_list = [
                    {
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 5,
                        "join_tolerance": 5,
                    },
                    {
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                    },
                    {
                        "vertical_strategy": "explicit",
                        "horizontal_strategy": "explicit",
                    },
                    {
                        "vertical_strategy": "lines_strict",
                        "horizontal_strategy": "lines_strict",
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                    }
                ]
                
                for settings_idx, table_settings in enumerate(table_settings_list):
                    try:
                        tables = page.extract_tables(table_settings=table_settings)
                        
                        if tables:
                            print(f"    策略 {settings_idx + 1} 找到 {len(tables)} 个表格")
                            
                            for table_num, table in enumerate(tables):
                                if not table or len(table) == 0:
                                    continue
                                
                                # 检查表格内容
                                if len(table) < 2:
                                    continue
                                
                                # 将整个表格转换为文本以便搜索
                                all_table_text = ' '.join([
                                    ' '.join([str(cell).strip() if cell else '' for cell in row])
                                    for row in table
                                ])
                                
                                # 检查是否包含目标关键词
                                has_particle = 'Particle' in all_table_text or 'particle' in all_table_text.lower()
                                has_size = 'Size' in all_table_text or 'size' in all_table_text.lower()
                                has_cumulative = 'Cumulative' in all_table_text or 'cumulative' in all_table_text.lower()
                                has_counts = 'Counts' in all_table_text or 'counts' in all_table_text.lower()
                                
                                if has_particle and has_size and has_cumulative and has_counts:
                                    print(f"    找到包含目标关键词的表格！")
                                    
                                    # 转换为DataFrame
                                    df = pd.DataFrame(table[1:], columns=table[0])
                                    
                                    # 移除完全为空的行和列
                                    df = df.dropna(how='all').dropna(axis=1, how='all')
                                    
                                    print(f"    表格行数: {len(df)}, 列数: {len(df.columns)}")
                                    # 强制打印所有列名
                                    col_names = [str(col) for col in df.columns]
                                    print(f"    表格列名: {col_names}")
                                    
                                    # 打印表头行以便调试
                                    if len(table) > 0:
                                        print(f"    表头行: {table[0]}")
                                    
                                    # 查找目标列 - 直接匹配已知的列名格式
                                    particle_size_col = None
                                    cumulative_counts_col = None
                                    
                                    for col in df.columns:
                                        col_str = str(col) if col else ''
                                        col_str_normalized = col_str.replace('\n', ' ').replace('\r', ' ')
                                        
                                        # 匹配 Particle Size 列
                                        if 'Particle Size' in col_str and ('µm' in col_str or 'um' in col_str):
                                            particle_size_col = col
                                        
                                        # 匹配 Cumulative Counts/mL 列（处理换行符）
                                        col_str_normalized_lower = col_str_normalized.lower()
                                        if 'cumulative' in col_str_normalized_lower and 'counts' in col_str_normalized_lower and '/ml' in col_str_normalized_lower:
                                            cumulative_counts_col = col
                                    
                                    # 如果还是没找到，尝试更宽松的匹配
                                    if not particle_size_col:
                                        for col in df.columns:
                                            col_str_lower = str(col).lower() if col else ''
                                            if 'particle' in col_str_lower and 'size' in col_str_lower:
                                                particle_size_col = col
                                                break
                                    
                                    if not cumulative_counts_col:
                                        for col in df.columns:
                                            col_str_normalized = str(col).replace('\n', ' ').replace('\r', ' ').lower() if col else ''
                                            if 'cumulative' in col_str_normalized and 'counts' in col_str_normalized:
                                                cumulative_counts_col = col
                                                break
                                    
                                    # 如果列名匹配不完整，使用列索引作为备用方案
                                    if not particle_size_col or not cumulative_counts_col:
                                        # 根据调试输出，列顺序是：['Run No.', 'Particle Size(µm)', 'Cumulative Count', 'Differential Count', 'Cumulative\nCounts/mL', 'Differential\nCounts/mL']
                                        # 所以 Particle Size 是第2列（索引1），Cumulative Counts/mL 是第5列（索引4）
                                        if len(df.columns) >= 5:
                                            try:
                                                if not particle_size_col:
                                                    particle_size_col = df.columns[1]  # Particle Size(µm)
                                                if not cumulative_counts_col:
                                                    cumulative_counts_col = df.columns[4]  # Cumulative\nCounts/mL
                                                print(f"    使用列索引: '{particle_size_col}' 和 '{cumulative_counts_col}'")
                                            except Exception as e:
                                                print(f"    使用列索引失败: {str(e)}")
                                    
                                    # 如果两个列都找到了，提取数据
                                    if particle_size_col and cumulative_counts_col:
                                        print(f"    准备提取数据，使用列: '{particle_size_col}' 和 '{cumulative_counts_col}'")
                                        try:
                                            # 提取数据 - 优先提取第21-25行
                                            if len(df) >= 25:
                                                extracted_data = df.iloc[19:24][[particle_size_col, cumulative_counts_col]].copy()
                                            elif len(df) >= 20:
                                                extracted_data = df.iloc[19:min(24, len(df)-1)][[particle_size_col, cumulative_counts_col]].copy()
                                            else:
                                                # 如果行数不足，提取所有可用行
                                                extracted_data = df[[particle_size_col, cumulative_counts_col]].copy()
                                            
                                            extracted_data.columns = ['Particle Size(µm)', 'Cumulative Counts/mL']
                                            extracted_data = extracted_data.dropna(how='all')
                                            
                                            # 转换数据类型
                                            try:
                                                extracted_data['Particle Size(µm)'] = pd.to_numeric(
                                                    extracted_data['Particle Size(µm)'], errors='coerce'
                                                )
                                                extracted_data['Cumulative Counts/mL'] = pd.to_numeric(
                                                    extracted_data['Cumulative Counts/mL'], errors='coerce'
                                                )
                                            except Exception as e:
                                                print(f"    数据类型转换警告: {str(e)}")
                                            
                                            extracted_data = extracted_data.dropna()
                                            
                                            # 如果提取的数据为空，尝试从整个表格中提取包含目标尺寸的行
                                            if extracted_data.empty:
                                                print(f"    第21-25行数据为空，尝试从整个表格提取目标尺寸数据")
                                                # 提取所有包含目标尺寸（2, 5, 10, 25, 50）的行
                                                target_sizes = [2, 5, 10, 25, 50]
                                                filtered_rows = []
                                                for _, row in df.iterrows():
                                                    size_val = row[particle_size_col]
                                                    if pd.notna(size_val):
                                                        try:
                                                            size_val = float(size_val)
                                                            if size_val in target_sizes:
                                                                filtered_rows.append(row[[particle_size_col, cumulative_counts_col]])
                                                        except:
                                                            pass
                                                
                                                if filtered_rows:
                                                    extracted_data = pd.DataFrame(filtered_rows)
                                                    extracted_data.columns = ['Particle Size(µm)', 'Cumulative Counts/mL']
                                                    extracted_data = extracted_data.dropna()
                                            
                                            if not extracted_data.empty:
                                                print(f"    成功提取 {len(extracted_data)} 行数据")
                                                return extracted_data
                                            else:
                                                print(f"    警告: 提取的数据为空")
                                        except Exception as e:
                                            print(f"    数据提取失败: {str(e)}")
                                            import traceback
                                            traceback.print_exc()
                    except Exception as e:
                        continue
                
                # 方法2: 如果表格提取失败，尝试直接从文本中提取
                # 提取页面文本
                page_text = page.extract_text()
                if page_text and ('Particle' in page_text or 'particle' in page_text.lower()):
                    print(f"    页面包含 'Particle' 关键词，尝试从文本中提取...")
                    # 这里可以添加文本解析逻辑，但先尝试表格提取
                    
    except Exception as e:
        print(f"  错误: 处理文件时出错 - {str(e)}")
        import traceback
        traceback.print_exc()
    
    print(f"  未能从文件中提取到数据")
    return None


def function1_extract_pdf():
    """功能1：提取PDF中样品数据"""
    print("\n" + "=" * 60)
    print("功能1：提取PDF中样品数据")
    print("=" * 60)
    
    # 1. 获取用户输入的PDF文件路径
    while True:
        pdf_path = input("\n请输入PDF文件所在文件夹路径: ").strip()
        
        # 去除引号（如果用户复制路径时带引号）
        pdf_path = pdf_path.strip('"').strip("'")
        
        if not pdf_path:
            print("错误: 路径不能为空，请重新输入")
            continue
        
        # 检查路径是否存在
        if not os.path.exists(pdf_path):
            print(f"错误: 路径不存在: {pdf_path}")
            continue
        
        if not os.path.isdir(pdf_path):
            print(f"错误: 不是有效的文件夹路径: {pdf_path}")
            continue
        
        break
    
    # 2. 查找所有PDF文件
    pdf_files = find_pdf_files(pdf_path)
    
    if not pdf_files:
        print(f"\n在路径 {pdf_path} 中未找到PDF文件！")
        input("\n按回车键返回...")
        return
    
    print(f"\n找到 {len(pdf_files)} 个PDF文件:")
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"  {i}. {os.path.basename(pdf_file)}")
    
    # 3. 处理每个PDF文件
    results = {}
    
    for pdf_file in pdf_files:
        file_name = os.path.basename(pdf_file)
        extracted_data = extract_table_from_pdf(pdf_file)
        
        if extracted_data is not None and not extracted_data.empty:
            results[file_name] = extracted_data
            print(f"\n[成功] 成功提取 {file_name} 的数据:")
            print(extracted_data.to_string(index=False))
        else:
            print(f"\n[失败] 未能从 {file_name} 提取数据")
    
    # 4. 将结果转换为汇总表格式
    if results:
        # 定义目标颗粒尺寸
        target_sizes = [2, 5, 10, 25, 50]
        
        # 创建汇总表
        summary_data = []
        
        for idx, (file_name, data) in enumerate(results.items(), 1):
            # 提取样品名称（去掉.pdf扩展名）
            sample_name = os.path.splitext(file_name)[0]
            
            # 创建一行数据
            row_data = {
                '序号': idx,
                '样品名称': sample_name
            }
            
            # 初始化所有颗粒尺寸列为0
            for size in target_sizes:
                row_data[f'≥{size} μm'] = 0
            
            # 从提取的数据中查找对应颗粒尺寸的值
            for _, row in data.iterrows():
                particle_size = row['Particle Size(µm)']
                cumulative_counts = row['Cumulative Counts/mL']
                
                # 检查是否是目标尺寸之一
                if pd.notna(particle_size) and pd.notna(cumulative_counts):
                    particle_size = float(particle_size)
                    cumulative_counts = float(cumulative_counts)
                    
                    # 找到匹配的尺寸列
                    for size in target_sizes:
                        if abs(particle_size - size) < 0.01:  # 允许小的浮点误差
                            row_data[f'≥{size} μm'] = cumulative_counts
                            break
            
            summary_data.append(row_data)
        
        # 创建汇总DataFrame
        summary_df = pd.DataFrame(summary_data)
        
        # 重新排列列的顺序
        column_order = ['序号', '样品名称'] + [f'≥{size} μm' for size in target_sizes]
        summary_df = summary_df[column_order]
        
        print(f"\n\n汇总表:")
        print(summary_df.to_string(index=False))
        
        # 5. 保存结果到Excel文件（只保存汇总表）
        # 输出文件保存到PDF文件所在的目录
        output_file = os.path.join(pdf_path, '提取结果.xlsx')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 只保存汇总表
            summary_df.to_excel(writer, sheet_name='结果汇总', index=False)
        
        print(f"\n\n所有结果已保存到: {output_file}")
        print(f"汇总表已保存到工作表: 结果汇总")
    else:
        print("\n未提取到任何数据，无法生成汇总表")
    
    print("\n" + "=" * 60)
    print("处理完成！")
    print("=" * 60)
    input("\n按回车键返回主菜单...")


def find_csv_files(directory):
    """
    查找指定目录下的所有CSV文件
    
    Args:
        directory: 要搜索的目录路径
        
    Returns:
        CSV文件路径列表
    """
    csv_files = []
    input_path = Path(directory)
    
    # 检查目录是否存在
    if not input_path.exists():
        print(f"错误: 路径不存在: {directory}")
        return []
    
    if not input_path.is_dir():
        print(f"错误: 不是有效的目录路径: {directory}")
        return []
    
    # 查找CSV文件
    for file in input_path.glob('*.csv'):
        csv_files.append(str(file))
    
    return sorted(csv_files)


def extract_csv_data(csv_path):
    """
    从CSV文件中提取数据（第1列和第5列的第31-42行）
    
    Args:
        csv_path: CSV文件路径
        
    Returns:
        提取的数据（DataFrame），如果失败则返回None
    """
    try:
        # 读取CSV文件
        df = pd.read_csv(csv_path, header=None)
        
        # 提取第31-42行（索引30-41）
        if len(df) < 42:
            print(f"  警告: CSV文件行数不足42行（实际{len(df)}行）")
            return None
        
        # 提取第1列（索引0）和第5列（索引4）的第31-42行
        extracted_data = df.iloc[30:42][[0, 4]].copy()
        extracted_data.columns = ['ESD类型', '数值']
        
        # 清理数据
        extracted_data = extracted_data.dropna()
        
        return extracted_data
        
    except Exception as e:
        print(f"  错误: 读取CSV文件失败 - {str(e)}")
        return None


def function2_extract_csv():
    """功能2：提取CSV中样本数据"""
    print("\n" + "=" * 60)
    print("功能2：提取CSV中样本数据")
    print("=" * 60)
    
    # 1. 获取用户输入的CSV文件路径
    while True:
        csv_path = input("\n请输入CSV文件所在文件夹路径: ").strip()
        
        # 去除引号（如果用户复制路径时带引号）
        csv_path = csv_path.strip('"').strip("'")
        
        if not csv_path:
            print("错误: 路径不能为空，请重新输入")
            continue
        
        # 检查路径是否存在
        if not os.path.exists(csv_path):
            print(f"错误: 路径不存在: {csv_path}")
            continue
        
        if not os.path.isdir(csv_path):
            print(f"错误: 不是有效的文件夹路径: {csv_path}")
            continue
        
        break
    
    # 2. 查找所有CSV文件
    csv_files = find_csv_files(csv_path)
    
    if not csv_files:
        print(f"\n在路径 {csv_path} 中未找到CSV文件！")
        input("\n按回车键返回...")
        return
    
    print(f"\n找到 {len(csv_files)} 个CSV文件:")
    for i, csv_file in enumerate(csv_files, 1):
        print(f"  {i}. {os.path.basename(csv_file)}")
    
    # 3. 处理每个CSV文件
    results = {}
    
    for csv_file in csv_files:
        file_name = os.path.basename(csv_file)
        print(f"\n正在处理: {file_name}")
        
        # 提取数据
        extracted_data = extract_csv_data(csv_file)
        
        if extracted_data is not None and not extracted_data.empty:
            # 处理文件名：去掉_summary后缀
            sample_name = os.path.splitext(file_name)[0]  # 去掉.csv
            if sample_name.endswith('_summary'):
                sample_name = sample_name[:-8]  # 去掉_summary
            
            results[sample_name] = extracted_data
            print(f"  [成功] 成功提取数据")
        else:
            print(f"  [失败] 未能提取数据")
    
    # 4. 将结果转换为汇总表格式
    if results:
        # 定义ESD类型列
        esd_columns = [
            'ESD 1-2 um', 'ESD 2-5 um', 'ESD 5-10 um', 'ESD 10-25 um', 
            'ESD 25-50 um', 'ESD 50 um+',
            'ESD 1-2 um SO', 'ESD 2-5 um SO', 'ESD 5-10 um SO', 
            'ESD 10-25 um SO', 'ESD 25-50 um SO', 'ESD 50 um +SO'
        ]
        
        # 创建汇总表
        summary_data = []
        
        for idx, (sample_name, data) in enumerate(results.items(), 1):
            # 创建一行数据
            row_data = {
                '序号': idx,
                '样品名称': sample_name
            }
            
            # 初始化所有ESD列为0
            for col in esd_columns:
                row_data[col] = 0
            
            # 从提取的数据中查找对应ESD类型的值
            # 按照顺序匹配：ESD 1-2 um, ESD 2-5 um, ESD 5-10 um, ESD 10-25 um, 
            # ESD 25-50 um, ESD 50 um+, ESD 1-2 um SO, ESD 2-5 um SO, 
            # ESD 5-10 um SO, ESD 10-25 um SO, ESD 25-50 um SO, ESD 50 um +SO
            for i, (_, row) in enumerate(data.iterrows()):
                esd_type = str(row['ESD类型']).strip()
                value = row['数值']
                
                if pd.notna(value) and i < len(esd_columns):
                    try:
                        value = float(value)
                        # 直接按顺序匹配（第31-42行对应12个ESD类型）
                        row_data[esd_columns[i]] = value
                    except:
                        pass
            
            summary_data.append(row_data)
        
        # 创建汇总DataFrame
        summary_df = pd.DataFrame(summary_data)
        
        # 重新排列列的顺序
        column_order = ['序号', '样品名称'] + esd_columns
        summary_df = summary_df[column_order]
        
        print(f"\n\n汇总表:")
        print(summary_df.to_string(index=False))
        
        # 5. 保存结果到Excel文件
        output_file = os.path.join(csv_path, '提取结果.xlsx')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='结果汇总', index=False)
        
        print(f"\n\n所有结果已保存到: {output_file}")
        print(f"汇总表已保存到工作表: 结果汇总")
    else:
        print("\n未提取到任何数据，无法生成汇总表")
    
    print("\n" + "=" * 60)
    print("处理完成！")
    print("=" * 60)
    input("\n按回车键返回主菜单...")


def show_menu():
    """显示主菜单"""
    print("\n" + "=" * 60)
    print("数据提取工具")
    print("=" * 60)
    print("\n请选择功能：")
    print("  1 - 提取PDF中样品数据")
    print("  2 - 提取CSV中样本数据")
    print("  ESC - 退出程序")
    print("\n" + "=" * 60)


def get_user_choice():
    """获取用户选择（支持数字键和ESC键）"""
    print("\n请输入选项（1/2）或按ESC退出: ", end='', flush=True)
    
    # 使用input方式（更兼容）
    try:
        choice = input().strip()
        if choice.lower() in ['esc', 'exit', 'q', 'quit']:
            return 'esc'
        return choice
    except (EOFError, KeyboardInterrupt):
        return 'esc'


def main():
    """主函数"""
    while True:
        show_menu()
        choice = get_user_choice()
        
        if choice == 'esc' or choice.lower() == 'esc':
            print("\n感谢使用，再见！")
            break
        elif choice == '1':
            function1_extract_pdf()
        elif choice == '2':
            function2_extract_csv()
        else:
            print("\n无效的选择，请重新输入！")
            input("按回车键继续...")


if __name__ == "__main__":
    main()

