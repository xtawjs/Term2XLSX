import sys
import re
import os
from openpyxl import Workbook
import subprocess
# import time


def process_file(lines):
    # 编译正则表达式
    pattern0 = re.compile(r"^\+[-+]+\+$")
    pattern1 = re.compile(r"^\|.+\|$")
    i = 0
    # 为了防止在遍历时修改列表长度，先记录需要删除的行号
    lines_to_delete = set()
    print('check for merging')
    while i < len(lines):
        if pattern0.match(lines[i]) and i > 0 and i + 3 < len(lines):
            # 检查前一行和后两行
            if pattern1.match(lines[i - 1]) and pattern1.match(
                    lines[i + 1]) and pattern0.match(lines[i + 2]) and pattern1.match(lines[i + 3]):
                # 添加当前行和后两行到删除列表
                lines_to_delete.update([i, i + 1, i + 2])
        i += 1

    # 从后向前删除行，这样不会影响前面行的索引
    for index in sorted(lines_to_delete, reverse=True):
        del lines[index]
    # print(lines)
    return lines


def extract_tables_according_to_pattern(file_lines):
    tables = []
    table_header = None
    table_rows = []
    header_detected = False
    content_detected = False

    pattern = re.compile(r"^\+[-+]+\+$")

    # 寻找所有匹配正则表达式的行的索引
    matched_lines = [index for index, line in enumerate(file_lines) if pattern.search(line)]

    # 计算匹配行的数量对3取余
    x = len(matched_lines) % 3
    # 根据x的值选择相应的处理方式
    if x == 1:
        # 选择第1个匹配行之后的所有内容
        start_index = matched_lines[0] + 1
        file_read = file_lines[start_index:]
    elif x == 2:
        # 选择第2个匹配行之后的所有内容
        if len(matched_lines) > 1:
            start_index = matched_lines[1] + 1
            file_read = file_lines[start_index:]
        else:
            # 如果没有足够的匹配行，返回空内容
            file_read = []
    else:
        # x == 0，使用整个文件内容
        file_read = file_lines

    for line in file_read:
        if re.match(pattern, line):
            if not header_detected:  # First occurrence, nothing to do yet
                header_detected = True
            elif header_detected and not content_detected:  # Second occurrence, start capturing rows
                content_detected = True
            elif header_detected and content_detected:  # Third occurrence, stop capturing, reset
                tables.append((table_header, table_rows))
                table_header = None
                table_rows = []
                header_detected = False
                content_detected = False
        else:
            if header_detected and not content_detected:
                table_header = line.strip()  # Capture the header
            elif header_detected and content_detected:
                table_rows.append(line.strip())  # Capture the content rows

    return tables


def save_tables_to_excel(tables, excel_path):
    wb = Workbook()
    ws = wb.active

    for index, (header, rows) in enumerate(tables):
        if index > 0:
            ws = wb.create_sheet(f'Sheet{index + 1}')

        columns = header.split('|')[1:-1]  # Ignore empty strings from first and last split
        columns = [col.strip() for col in columns]
        columns = [col.split(".")[-1] for col in columns]  # 去除表名
        ws.append(columns)

        for row in rows:
            data = row.split('|')[1:-1]
            data = [item.strip() for item in data]
            ws.append(data)
    wb.active = ws
    wb.save(excel_path)


def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <path_to_input_file>")
        sys.exit(1)

    input_file_path = sys.argv[1]

    if re.search(r'\.\w+$', input_file_path):
        excel_file_path = re.sub(r'\.\w+$', '.xlsx', input_file_path)
    else:
        # 如果没有匹配到扩展名，直接在文件名末尾添加新的扩展名
        excel_file_path = input_file_path + '.xlsx'

    # 尝试以 UTF-8 编码读取文件
    try:
        with open(input_file_path, 'r', encoding='utf-8') as file:
            file_content = process_file(file.readlines())
    # 如果出现解码错误，尝试以 GB18030 编码读取文件
    except UnicodeDecodeError:
        print('NOT UTF-8, USE GB18030')
        with open(input_file_path, 'r', encoding='gb18030') as file:
            file_content = process_file(file.readlines())

    print('transforming')
    extracted_tables = extract_tables_according_to_pattern(file_content)
    print('saving')
    save_tables_to_excel(extracted_tables, excel_file_path)
    print('All done')
    # time.sleep(1)
    # Open the generated Excel file automatically if possible
    if sys.platform.startswith('darwin'):
        subprocess.call(('open', excel_file_path))
    elif os.name == 'nt':  # For Windows
        os.startfile(excel_file_path)
    elif os.name == 'posix':  # For Linux variants
        subprocess.call(('xdg-open', excel_file_path))


if __name__ == "__main__":
    main()
