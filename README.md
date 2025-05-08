# 文档格式转换工具

一个强大的文档格式转换工具，支持Word、PDF和Markdown之间的相互转换。

## 功能

- **Word/PDF转Markdown**:
  - 支持将Word文档转换为Markdown
  - 支持将PDF文档转换为Markdown
  - 可选"简单转换"（一个文档对应一个Markdown文件）
  - 可选"分割转换"（按一级标题将Word分割为多个Markdown文件）
  - 支持多个文档合并为一个Markdown文件

- **Markdown转Word/PDF**:
  - 支持将Markdown文档转换为Word
  - 支持将Markdown文档转换为PDF
  - 支持多个Markdown文件合并为一个输出文件

## 依赖库

```
docx (python-docx)
pdfplumber
markdown
docxtpl
reportlab
pypandoc
PyQt5
```

## 安装依赖

```bash
pip install python-docx pdfplumber markdown docxtpl reportlab pypandoc PyQt5
```

## 使用方法

1. 运行主程序：

```bash
python main.py
```

2. 在打开的界面中选择所需的转换类型（"转换为Markdown"或"从Markdown转换"）

3. 根据提示选择：
   - 文件类型/目标格式
   - 要转换的文件
   - 转换模式（仅Word转Markdown时可用）
   - 是否合并多个文件
   - 输出目录

4. 点击"开始转换"按钮开始处理

## 文件结构

- `main.py`: 主程序入口
- `utils.py`: 工具函数模块，包含通用的转换功能
- `converters.py`: 转换器模块，包含转换线程类
- `word_to_md_combined_refactored.py`: 主界面和应用程序逻辑

## 注意事项

- 某些复杂格式（特别是复杂表格和嵌套格式）的转换可能不完美
- PDF转换依赖于PDF文档的内部结构，不同的PDF生成方式可能导致转换质量差异
- 如需使用合并功能转换多个文件，建议选择相似结构的文档

## 许可证

MIT 