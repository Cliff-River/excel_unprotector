import openpyxl

def remove_sheet_protection(file_path, output_path):
    # 加载工作簿
    wb = openpyxl.load_workbook(file_path)
    
    # 遍历所有工作表并取消保护
    for sheet in wb.worksheets:
        sheet.protection.disable()  # 禁用保护
    
    # 保存为新文件
    wb.save(output_path)
    print(f"破解完成！已保存至: {output_path}")

# 使用示例
remove_sheet_protection('data/protected_file.xlsx', 'data/unprotected_file.xlsx')