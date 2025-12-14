"""
PDF表格提取程序
功能：
1. 读取当前目录下INPUT文件夹内的PDF文件
2. 解析PDF文件内容，提取表格中特定字段和值
3. 将结果汇总并保存到Excel文件
"""

import os
import pdfplumber
import pandas as pd
from pathlib import Path


def find_pdf_files(directory='INPUT'):
    """
    查找指定目录下的所有PDF文件
    
    Args:
        directory: 要搜索的目录，默认为INPUT文件夹  
        
    Returns:
        PDF文件路径列表
    """
    pdf_files = []
    input_path = Path(directory)
    
    # 检查INPUT文件夹是否存在
    if not input_path.exists():
        print(f"警告: {directory} 文件夹不存在，将创建该文件夹")
        input_path.mkdir(parents=True, exist_ok=True)
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


def main():
    """
    主函数
    """
    print("=" * 60)
    print("PDF表格提取程序")
    print("=" * 60)
    
    # 1. 查找所有PDF文件
    pdf_files = find_pdf_files()
    
    if not pdf_files:
        print("\n未找到PDF文件！")
        return
    
    print(f"\n找到 {len(pdf_files)} 个PDF文件:")
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"  {i}. {os.path.basename(pdf_file)}")
    
    # 2. 处理每个PDF文件
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
    
    # 3. 将结果转换为汇总表格式
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
        
        # 4. 保存结果到Excel文件（只保存汇总表）
        output_file = '提取结果.xlsx'
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 只保存汇总表
            summary_df.to_excel(writer, sheet_name='结果汇总', index=False)
        
        print(f"\n\n所有结果已保存到: {output_file}")
        print(f"汇总表已保存到工作表: 结果汇总")
    
    print("\n" + "=" * 60)
    print("处理完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()

