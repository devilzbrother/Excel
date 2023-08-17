#Remember to replace 'path_to_existing_excel.xlsx' and 'path_to_new_excel.xlsx' with the actual file paths.

# Read data from existing Excel file
read_data = ExcelHandler.read_excel('path_to_existing_excel.xlsx')

# Write data to a new Excel file
ExcelHandler.write_excel(read_data, 'path_to_new_excel.xlsx')