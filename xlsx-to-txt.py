import openpyxl
import os

# Path to the Excel file
excel_path = 'path_to_your_excel_file.xlsx'

# Load the workbook and select the worksheet you're working on
workbook = openpyxl.load_workbook(excel_path)
sheet = workbook.active

# Directory where you want the .txt files to be saved
output_directory = 'path_to_output_directory'
os.makedirs(output_directory, exist_ok=True)

# Specify the column whose cells you want to extract (e.g., 'A')
target_column = 'A'

# Base word for the naming convention of .txt files
base_word = 'filename'

# Extract non-empty cells from the specified column
cells_in_column = [cell for cell in sheet[target_column] if cell.value is not None]

# Loop through each cell and save its content as a .txt file
for index, cell in enumerate(cells_in_column, start=1):
    # Generate the file name using the format: filename_1, filename_2, etc.
    file_name = f'{base_word}_{index}.txt'
    file_path = os.path.join(output_directory, file_name)
    
    # Write the cell's content to the .txt file
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(str(cell.value))

print("Process complete")
