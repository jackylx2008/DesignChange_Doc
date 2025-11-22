# Docx_file_fulfill

## 项目简介

用于自动填充设计变更审批表的工具，从 Excel 文件读取数据，基于 Word 模板生成批量审批文档。

## 环境要求

- Python 3.12+
- 推荐使用虚拟环境

## 依赖包

项目依赖以下 Python 包：

```txt
numpy<2              # 数组计算库（需使用 1.x 版本以避免兼容性问题）
pandas>=2.0.0        # 数据处理和分析
python-docx>=1.0.0   # Word 文档操作
openpyxl>=3.0.0      # Excel 文件读写
```

## 安装步骤

1. 克隆项目：

```bash
git clone https://github.com/jackylx2008/DesignChange_Doc.git
cd DesignChange_Doc
```

2. 创建虚拟环境（推荐）：

```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1  # Windows PowerShell
```

3. 安装依赖：

```bash
pip install -r requirements.txt
```

## 使用说明

运行主程序：

```bash
python Docx_fulfill/DesignChange_Doc.py
```

## 注意事项

- 本程序仅支持 Windows 平台
- Excel 文件中的单元格必须填满，否则会报错
- 需要 NumPy 1.x 版本，不兼容 NumPy 2.x
