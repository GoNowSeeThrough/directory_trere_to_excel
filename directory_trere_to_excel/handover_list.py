#!/usr/bin/env python3

# To package this script for distribution:
# 1. Install pyinstaller: pip install pyinstaller
# 2. Run the tests: python test_directory/test_handover_list.py
# 3. For Windows: pyinstaller --onefile --additional-hooks-dir=. handover_list.py
# 4. For macOS: pyinstaller --onefile --windowed --additional-hooks-dir=. handover_list.py

import os
import datetime

try:
    import xlsxwriter
except ImportError:
    print("Error: xlsxwriter module not found.")
    print("Please install it using: pip install xlsxwriter")
    exit()

def generate_handover_list():
    """Generates an Excel handover list based on the current working directory."""

    current_directory = os.getcwd()
    today_date = datetime.date.today().strftime("%Y-%m-%d")
    excel_filename = f"交接清单{today_date}.xlsx"

    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet = workbook.add_worksheet()

    # Write headers
    headers = ["项目名称", "内容模块", "文件名", "公司电脑位置", "交接人"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    row = 1  # Start writing data from the second row
    for dirpath, dirnames, filenames in os.walk(current_directory):
        # Project Name: Parent folder name
        project_name = os.path.basename(current_directory)

        # Content Module: Subfolder name (relative to current_directory)
        if dirpath == current_directory:
            content_module = ""  # Or use a placeholder like "Root"
        else:
            content_module = os.path.relpath(dirpath, current_directory)

        # Handle files in the current directory
        for filename in filenames:
            # File Name
            file_name = filename

            # Company Computer Location: Absolute file path
            absolute_file_path = os.path.abspath(os.path.join(dirpath, filename))

            worksheet.write(row, 0, project_name)
            worksheet.write(row, 1, content_module)
            worksheet.write(row, 2, file_name)
            worksheet.write(row, 3, absolute_file_path)
            worksheet.write(row, 4, "")  # Empty "交接人" column

            row += 1

        # Handle empty subfolders
        if not filenames and dirpath != current_directory: #check to see it is not the root folder
            worksheet.write(row, 0, project_name)
            worksheet.write(row, 1, content_module)
            worksheet.write(row, 2, "")
            worksheet.write(row, 3, "")
            worksheet.write(row, 4, "")
            row+=1


    workbook.close()
    print(f"Excel handover list created: {excel_filename}")


import logging

logging.basicConfig(filename='error.log', level=logging.ERROR)

if __name__ == "__main__":
    try:
        generate_handover_list()
    except Exception as e:
        logging.exception("An error occurred:")
        print(f"An error occurred: {e}. See error.log for details.")
