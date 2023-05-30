import os
import openpyxl
import time

def update_excel(file_path):
    # Extract the relevant information from the file name
    file_name = os.path.basename(file_path)
    content = file_name.split('.')[0]  # Assuming the content is before the file extension

    # Define the conditions and their respective positions
    conditions = {
        'ISA00045-01': (1, 2),
        'ISA00045-02': (2, 2),
        'ISA00045-03': (3, 2),
        # Add more conditions as needed
    }


        # Attempt to open the Excel spreadsheet with retry logic

    max_retries = 999

    retries = 0

    while retries < max_retries:

        try:

            workbook = openpyxl.load_workbook('your_spreadsheet.xlsx')

            break

        except Exception as e:

            print(f"Error opening the Excel file: {e}")

            print("Retrying...")

            retries += 1

            time.sleep(15)  # Delay before retrying
    
    # Check if the condition exists in the file name and the file content contains "PASS"
    for condition, position in conditions.items():
        if condition in content and 'PASS' in open(file_path).read():
            row, column = position
            current_value = worksheet.cell(row=row, column=column).value
            new_value = current_value + 1 if current_value else 1
            worksheet.cell(row=row, column=column, value=new_value)
            break  # Exit the loop if condition is found

    # Save the changes to the spreadsheet
    workbook.save('data.xlsx')

# Recursive function to scan all folders within the directory for new text files
def scan_directory(directory_path):
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith('.txt'):
                file_path = os.path.join(root, file)
                modified_time = os.path.getmtime(file_path)
                if modified_time > start_time and file_path not in processed_files:
                    update_excel(file_path)
                    processed_files.add(file_path)
                    print(f"Processed {file_path}.")

# Monitor a directory for new text files
directory_path = 'text_files'
start_time = time.time()
processed_files = set()

while True:
    scan_directory(directory_path)

    # Add a delay before checking for new files again
    time.sleep(1)  # Delay of 1 second
