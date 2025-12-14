# PDF表格提取程序

这个程序用于从PDF文件中提取特定表格数据。

## 功能

1. **自动扫描当前目录下的所有PDF文件**
2. **识别包含特定表头的表格**：
   - Run No.
   - Particle Size(µm)
   - Cumulative Count
   - Differential Count
   - Cumulative Counts/mL
   - Differential Counts/mL
3. **提取指定列的第21-25行数据**：
   - Particle Size(µm)
   - Cumulative Counts/mL
4. **将结果保存到Excel文件**

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

1. 将PDF文件放在程序所在目录
2. 运行程序：
```bash
python extract_pdf_tables.py
```

## 输出

- 程序会在控制台显示提取的数据
- 所有结果会保存到 `提取结果.xlsx` 文件中，每个PDF文件对应一个工作表

## 示例输出

对于文件 `HLX05-OT3-251203-U-L-25-1D.pdf`，提取的数据格式如下：

| Particle Size(µm) | Cumulative Counts/mL |
|-------------------|---------------------|
| 2.000             | 8.33                |
| 5.000             | 1.00                |
| 10.000            | 0.00                |
| 25.000            | 0.00                |
| 50.000            | 0.00                |


