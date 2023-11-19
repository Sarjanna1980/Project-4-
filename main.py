import openpyxl
from openpyxl import load_workbook
from openpyxl import load_workbook

def copy_values_between_files(source_file, source_sheet, start_row, end_row, source_column, target_file, target_sheet, target_column):
    source_workbook = load_workbook(source_file)
    source_sheet = source_workbook[source_sheet]

    target_workbook = load_workbook(target_file)
    target_sheet = target_workbook[target_sheet]

    target_row = 27  # Start copying to P27

    for row_index in range(start_row, end_row + 1, step):
        source_cell = source_sheet.cell(row=row_index, column=source_column)
        target_cell = target_sheet.cell(row=target_row, column=target_column)
        target_cell.value = source_cell.value

        target_row += 1

    target_workbook.save(target_file)
    source_workbook.close()
    target_workbook.close()

# Example usage:
source_file ="C:/Users/Sahar/Desktop/leonid_test/Test data.xlsx"
source_sheet = "sheet2"
start_row = 5
end_row = 28
source_column = 3  # Column C
step = 2
target_file ="C:/Users/Sahar/Desktop/leonid_test/UpdatedFormatForDCIRRPT.xlsx"
target_sheet = "Sheet1"
target_column = 16  # Column P
print(target_sheet)
copy_values_between_files(source_file, source_sheet, start_row, end_row, source_column, target_file, target_sheet, target_column)




#print_values(file_path, sheet_name, start_row, end_row, column, step)

#"C:/Users/Sahar/Desktop/leonid_test/Test data.xlsx"
#"C:/Users/Sahar/Desktop/leonid_test/Format for DCIR RPT.xlsx"