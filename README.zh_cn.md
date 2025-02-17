# README

[English](./README.md) | [中文](./README.zh_cn.md)

## Term2XLSX

这个 Python 脚本从文本文件中提取表格并将其保存到 Excel 文件中。它专门处理包含特定格式表格的文本文件（例如使用 `+` 和 `|` 字符）。脚本处理文件，提取表格，并将其保存到包含多个工作表的 Excel 工作簿中。适用于数据库在终端中的临时SQL查询结果转换为xlsx表格。

### 功能

- **表格提取**：根据特定模式从文本文件中提取表格。

- **多工作表支持**：每个表格保存到 Excel 文件中的单独工作表。

- **编码支持**：处理 UTF-8 和 GB18030 (兼容 GBK GB2312) 编码的文件。

- **自动打开文件**：处理完成后自动打开生成的 Excel 文件（支持 macOS、Windows 和 Linux）。

### 使用方法

1. **前提条件**：
   
   - Python 3.x
   
   - `openpyxl` 库（通过 `pip install openpyxl` 安装）

2. **运行脚本**：
   
   - 从命令行运行脚本：
     
     ```bash
     python tmp2xlsx.py <输入文件路径>
     ```
   
   - 将 `<输入文件路径>` 替换为你的文本文件路径。

3. **输出**：
   
   - 脚本将生成一个与输入文件同名的 Excel 文件，但扩展名为 `.xlsx`。
   
   - 文本文件中的每个表格将保存到 Excel 文件中的单独工作表。

### 示例

给定一个文本文件 `example.txt`，内容如下：

```
+---------+---------+
| 列1     | 列2     |
+---------+---------+
| 数据1   | 数据2   |
| 数据3   | 数据4   |
+---------+---------+
```

运行脚本：

```bash
python tmp2xlsx.py example.txt
```

将生成一个 Excel 文件 `example.xlsx`，其中包含表格的工作表。

**终端工具中的便捷用法：**

如Xshell “工具”-“选项”-“高级”-“文本编辑器”设置为该程序（pyinstaller打包时建议不要使用-F参数，影响启动速度）。设置后点击“编辑”-“要tmp2xlsx”-“全部”即可一键导出终端文本（滚动缓冲区）为xlsx文件。

### 注意事项

- 脚本假设表格使用 `+` 和 `|` 字符格式化。

- 如果文件编码不是 UTF-8，脚本将尝试以 GB18030 (兼容 GBK GB2312) 编码读取。

## Third-Party Licenses

This project uses the following third-party libraries:

- **openpyxl** (MIT License): https://openpyxl.readthedocs.io
