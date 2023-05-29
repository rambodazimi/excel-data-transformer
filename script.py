import pandas as pd
import os
import subprocess

try:
   import fsspec
except ImportError:
    print("The 'fsspec' library is not found. Installing it...")
    subprocess.check_call(['pip', 'install', 'fsspec'])

try:
    # Specify the file path of the Excel file
    file_path = 'C://temp//123.xlsx'

    # Check if the Excel file exists
    if not os.path.isfile(file_path):
        raise FileNotFoundError("The Excel file does not exist.")

    # Read the Excel file from the specified path
    data_frame = pd.read_excel(file_path)

    # Get the 7th column (column index 6) from the excel file and save the contents to column_7 variable
    column_7 = data_frame.iloc[:, 6]

    # Remove the 11th character from each element in column 7 and save them in column_7_updated variable
    column_7_updated = column_7.str.slice(0, 10) + column_7.str.slice(11)

    # Assign the updated column 7 back to the DataFrame
    data_frame.iloc[:, 6] = column_7_updated

    # Extract the directory path of the Excel file
    directory_path = os.path.dirname(file_path)

    # Construct the file path for the CSV file in the same directory that the excel file exists
    csv_file_path = os.path.join(directory_path, 'final.csv')

    # Save the modified DataFrame to the CSV file
    data_frame.to_csv(csv_file_path, index=False)

    print("CSV file saved successfully.")
    
except FileNotFoundError as e:
    print(f"Error: {e}")
    
except Exception as e:
    print(f"An error occurred: {e}")