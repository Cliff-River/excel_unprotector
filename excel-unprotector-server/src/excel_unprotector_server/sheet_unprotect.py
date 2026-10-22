import io
import openpyxl
from os import path

def remove_sheet_protection(file_path, output_path):
    wb = openpyxl.load_workbook(file_path)
    for sheet in wb.worksheets:
        sheet.protection.disable()
    wb.save(output_path)
    print(f"破解完成！已保存至: {output_path}")

def remove_sheet_protection_stream(input_stream):
    wb = openpyxl.load_workbook(input_stream)
    for sheet in wb.worksheets:
        sheet.protection.disable()
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream

if __name__ == "__main__":
    name = "protected_file.xlsx"
    # 使用示例
    remove_sheet_protection(path.join('data', name), path.join('data', f"unprotected_{name}"))
