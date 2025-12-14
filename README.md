# PDF表格提取工具

## 项目结构

```
readPDF/
├── src/                    # 源代码目录
│   ├── extract_pdf_tables.py    # 主程序源代码
│   ├── build_exe.py             # 打包脚本
│   └── requirements.txt         # Python依赖包
│
├── release/                # 发布文件目录（可分发）
│   ├── PDF表格提取工具.exe      # 可执行程序
│   └── 使用说明.txt             # 用户使用说明
│
├── docs/                   # 文档目录
│   ├── README.md           # 开发文档
│   ├── 打包说明.md         # 打包说明
│   └── 使用说明.txt        # 使用说明（副本）
│
├── build/                  # 构建临时文件（可删除）
├── dist/                   # PyInstaller输出（可删除）
├── INPUT/                  # 测试PDF文件（开发用）
└── README.md              # 本文件
```

## 快速开始

### 开发环境

1. 安装依赖：
   ```bash
   cd src
   pip install -r requirements.txt
   ```

2. 运行程序：
   ```bash
   python extract_pdf_tables.py
   ```

3. 打包程序：
   ```bash
   python build_exe.py
   ```

### 分发程序

从 `release/` 目录获取以下文件分发给用户：
- `PDF表格提取工具.exe`
- `使用说明.txt`

## 目录说明

- **src/**: 源代码和开发相关文件
- **release/**: 可分发给用户的文件
- **docs/**: 项目文档
- **build/**: PyInstaller构建临时文件（可删除）
- **dist/**: PyInstaller输出目录（可删除）
- **INPUT/**: 测试用的PDF文件（开发用）

## 注意事项

- `build/` 和 `dist/` 目录是打包时自动生成的，可以删除
- 打包后，exe文件会自动复制到 `release/` 目录
- 源代码修改后需要重新打包才能更新exe文件
