import openpyxl
from openpyxl import load_workbook
from openpyxl import load_workbook

def copy_values_v1_between_files(source_file, source_sheet, start_row, end_row, source_column, target_file, target_sheet, target_column):
    source_workbook = load_workbook(source_file)
    source_sheet = source_workbook[source_sheet]

    target_workbook = load_workbook(target_file)
    target_sheet = target_workbook[target_sheet]

    target_row = 27  # Start copying to P27

    for row_index in range(start_row, end_row + 1, step):
        source_cell_v1 = source_sheet.cell(row=row_index, column=source_column_v1)
        target_cell_v1 = target_sheet.cell(row=target_row, column=target_column_v1)
        source_cell_v2 = source_sheet.cell(row=row_index, column=source_column_v2)
        target_cell_v2 = target_sheet.cell(row=target_row, column=target_column_v2)
        target_cell_v1.value = source_cell_v1.value
        target_cell_v2.value = source_cell_v2.value

        target_row += 1

    target_workbook.save(target_file)
    source_workbook.close()
    target_workbook.close()



# Example usage:
source_file ="C:/Users/Sahar/Desktop/leonid_test/Test data.xlsx"
source_sheet = "sheet2"
start_row = 6
end_row = 28
source_column_v1 = 3  # Column C
source_column_v2 = 4
step = 2
target_file ="C:/Users/Sahar/Desktop/leonid_test/UpdatedFormatForDCIRRPT.xlsx"
target_sheet = "Sheet1"
target_column_v1 = 16  # Column P
target_column_v2= 17
copy_values_v1_between_files(source_file, source_sheet, start_row, end_row, source_column_v1, target_file, target_sheet, target_column_v1)




#print_values(file_path, sheet_name, start_row, end_row, column, step)

#"C:/Users/Sahar/Desktop/leonid_test/Test data.xlsx"
#"C:/Users/Sahar/Desktop/leonid_test/Format for DCIR RPT.xlsx"