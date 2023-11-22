from openpyxl import load_workbook
import customtkinter
from tkinter import filedialog
import os
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

root = customtkinter.CTk()
root.title("SolarEdge")
root.iconbitmap("C:/Users/Sahar/Desktop/leonid_test/SEDG.ICO")
root.geometry(f"{900}x{600}")
root.grid_columnconfigure((0, 1, 2, 3), weight=1)
root.grid_rowconfigure((0, 1, 2, 3, 5, 6, 7, 8, 9, 10, 11), weight=1)

# Declare global variables
source_file_paths = [""] * 10
target_file_path = ""  # Initialize target_file_path

def process():
    # Adjust these values based on your requirements
    source_config = {
        'sheet': 'sheet2',
        'start_row': 6,
        'end_row': 28,
        'column_v1': 3,
        'column_v2': 4,
        'step': 2,
    }

    for i in range(10):
        entry_path_value = entry_paths[i].get()
        if entry_path_value:
            # Create a new target configuration for each iteration
            target_config = {
                'sheet': 'SOC X%',
                'target_column_v1': 16 + i * 6,  # Adjust this calculation based on your needs
                'target_column_v2': 17 + i * 6,  # Adjust this calculation based on your needs
                'target_row': 27,
            }
            copy_values_between_files(
                source_file_paths[i], source_config,
                target_file_path, target_config
            )

    root.destroy()

def choose_source_file(i):
    global source_file_paths
    file_path = filedialog.askopenfilename(title=f"Select Source File {i + 1}")
    entry_paths[i].delete(0, customtkinter.END)
    entry_paths[i].insert(0, file_path)
    source_file_paths[i] = file_path

# Create entry widgets and buttons for source file paths
entry_paths = []
buttons_choose_source = []

for i in range(10):
    entry_path = customtkinter.CTkEntry(root, width=330)
    entry_path.grid(row=i, column=0, padx=1, pady=1)

    button_choose_source = customtkinter.CTkButton(root, text=f"Choose Source File {i + 1}",
                                                    command=lambda i=i: choose_source_file(i))
    button_choose_source.grid(row=i, column=1, padx=1, pady=1)

    entry_paths.append(entry_path)
    buttons_choose_source.append(button_choose_source)

button_process = customtkinter.CTkButton(root, text="Process File", command=process)
button_process.grid(row=10, column=3, columnspan=2, padx=10, pady=10)

def copy_values_between_files(source_file, source_config, target_file, target_config):
    source_workbook = load_workbook(source_file)
    source_sheet = source_workbook[source_config['sheet']]

    target_workbook = load_workbook(target_file)
    target_sheet = target_workbook[target_config['sheet']]

    for row_index in range(source_config['start_row'], source_config['end_row'] + 1, source_config['step']):
        source_cell_v1 = source_sheet.cell(row=row_index, column=source_config['column_v1'])
        target_cell_v1 = target_sheet.cell(row=target_config['target_row'], column=target_config['target_column_v1'])
        source_cell_v2 = source_sheet.cell(row=row_index, column=source_config['column_v2'])
        target_cell_v2 = target_sheet.cell(row=target_config['target_row'], column=target_config['target_column_v2'])
        target_cell_v1.value = source_cell_v1.value
        target_cell_v2.value = source_cell_v2.value

        target_config['target_row'] += 1

    # Generate the updated target file name
    updated_target_file_path = generate_updated_filename(target_file)

    # Save the target workbook with the updated name
    target_workbook.save(updated_target_file_path)
    source_workbook.close()
    target_workbook.close()

# Function to generate the updated file name
def generate_updated_filename(original_filepath):
    if not original_filepath:
        return ""  # Return an empty string if the original filepath is empty

    filepath, filename = os.path.split(original_filepath)
    filename, extension = os.path.splitext(filename)
    updated_filename = f"{filename}_Updated{extension}"
    return os.path.join(filepath, updated_filename)

# Initialize target file information
target_file_path = "C:/Users/Sahar/Desktop/leonid_test/FormatForDCIRRPT.xlsx"

# Set target sheet information (SOC X%)
target_sheet = "SOC X%"
target_column_v1 = 16
target_column_v2 = 17
target_row = 27

root.mainloop()

