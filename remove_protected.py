# import os

# def remove_read_only_attribute(directory):
#     for filename in os.listdir(directory):
#         if filename.endswith(".xlsx"):
#             file_path = os.path.join(directory, filename)
            
#             # Remove the read-only attribute if set
#             os.chmod(file_path, 0o777)  # Grant read, write, and execute permissions to everyone
            
#             print(f"Enabled editing for {filename}")

# # Define the directory containing the Excel files
#   # Change this to your directory path

# # Run the function
# remove_read_only_attribute(directory)

import os
import openpyxl

def enable_editing_in_excel_files(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(directory, filename)
            
            try:
                # Load the workbook
                workbook = openpyxl.load_workbook(file_path)
                
                # Save the workbook again to ensure it's editable
                workbook.save(file_path)
                
                print(f"Enabled editing for {filename}")
            
            except Exception as e:
                print(f"Error processing {filename}: {e}")


# Define the directory containing the Excel files
directory = 'C:/Users/BRSBAWAM/Desktop/mapweb'  # Change this to your directory path

# Run the function
enable_editing_in_excel_files(directory)
