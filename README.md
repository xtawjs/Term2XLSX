# README

[English](./README.md) | [中文](./README.zh_cn.md)

## Term2XLSX

This Python script extracts tables from a text file and saves them into an Excel file. It is designed to handle text files containing tables formatted with specific patterns (e.g., using `+` and `|` characters). The script processes the file, extracts the tables, and saves them into an Excel workbook with multiple sheets.It is suitable for converting the results of temporary SQL queries in the terminal of the database to xlsx tables.

### Features

- **Table Extraction**: Extracts tables from text files based on specific patterns.

- **Multi-Sheet Support**: Each table is saved into a separate sheet in the Excel file.

- **Encoding Support**: Handles both UTF-8 and GB18030 encoded files.

- **Automatic File Opening**: Automatically opens the generated Excel file after processing (supported on macOS, Windows, and Linux).

### Usage

1. **Prerequisites**:
   
   - Python 3.x
   
   - `openpyxl` library (install via `pip install openpyxl`)

2. **Running the Script**:
   
   - Run the script from the command line:
     
     ```bash
     python script.py <path_to_input_file>
     ```
   
   - Replace `<path_to_input_file>` with the path to your text file.

3. **Output**:
   
   - The script will generate an Excel file with the same name as the input file but with a `.xlsx` extension.
   
   - Each table in the text file will be saved into a separate sheet in the Excel file.

### Example

Given a text file `example.txt` with the following content:

```
+---------+---------+
| Column1 | Column2 |
+---------+---------+
| Data1   | Data2   |
| Data3   | Data4   |
+---------+---------+
```

Running the script:

```bash
python script.py example.txt
```

Will generate an Excel file `example.xlsx` with a sheet containing the table.

**Convenient Usage of tmp2xlsx in Terminal Tools:**

For example, in Xshell, navigate to "Tools" -> "Options" -> "Advanced" -> "Text Editor" and set it to this program (when packaging with pyinstaller, it is recommended not to use the -F parameter, as it may affect startup speed). After setting, click "Edit" -> "To tmp2xlsx" -> "All" to export the terminal text (scroll buffer) to an xlsx file with one click. 

### Notes

- The script assumes that tables are formatted with `+` and `|` characters.

- If the file encoding is not UTF-8, the script will attempt to read it as GB18030.

## Third-Party Licenses

This project uses the following third-party libraries:

- **openpyxl** (MIT License): https://openpyxl.readthedocs.io
