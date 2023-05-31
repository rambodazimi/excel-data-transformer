import pandas as pd
import os
import subprocess


# copy the file specified in csv_file_path and add 1 at the end of its name
def copy_csv_file_with_suffix(csv_file_path, suffix):
    folder_path, file_name = os.path.split(csv_file_path)
    new_file_name = f"{file_name.split('.')[0]}{suffix}.{file_name.split('.')[1]}"
    new_file_path = os.path.join(folder_path, new_file_name)
    with open(csv_file_path, 'rb') as file_in, open(new_file_path, 'wb') as file_out:
        file_out.write(file_in.read())
    return new_file_path


def cut_excel_file(source_file_path, destination_folder_path):
    file_name = os.path.basename(source_file_path)
    destination_path = os.path.join(destination_folder_path, file_name)
    os.rename(source_file_path, destination_path)
    return destination_path

try:
   import fsspec
except ImportError:
    print("The 'fsspec' library is not found. Installing it...")
    subprocess.check_call(['pip', 'install', 'fsspec'])

try:
    print("Starting the script...")
    # Specify the folder path where the original Excel file exists
    folder_path = 'C://temp//Trimuph//Upload'

    found_excel_file = False
    print("Searching for the Excel file in the Upload folder...")
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls"):
                found_excel_file = True
                file_path = os.path.join(root, file)
                print("Found the Excel file!")
                break

    if(not found_excel_file):
        raise FileNotFoundError("The Excel file does not exist.")


    print("Processing the Excel file...")

    # Read the Excel file from the specified path
    data_frame = pd.read_excel(file_path)

    # Get the 24th column (column X) from the excel file and save the contents
    column_24 = data_frame.iloc[:, 23]

    # Remove the 11th character from each element in column X and save them
    column_24_updated = column_24.str.slice(0, 10) + column_24.str.slice(11)

    # Assign the updated column X back to the DataFrame
    data_frame.iloc[:, 23] = column_24_updated

    # Extract the directory path of the Excel file
    directory_path = 'C://temp//Trimuph'

    print("Creating a CSV file...")

    # Construct the file path for the CSV file
    csv_file_path = os.path.join(directory_path, 'tms_shipments.csv')

    # Save the modified DataFrame to the CSV file
    data_frame.to_csv(csv_file_path, index=False)

    print("CSV file saved successfully!")

    print("Making a copy of the CSV file...")
    my_path = 'C://temp//Trimuph//tms_shipments.csv'
    suffix = "1"

    copy_csv_file_with_suffix(my_path, suffix)

    print("Created tms_shipments1.csv successfully!")


    print("Moving the Excel file to the History folder...")

    source_file_path = file_path
    destination_folder_path = "C://temp//Trimuph//History"
    destination_file_path = cut_excel_file(source_file_path, destination_folder_path)
    print("Successfully moved the original excel file to the History folder")


    print("100% SUCCESSFUL!")


except FileNotFoundError as e:
    print(f"Error: {e}")
    
except Exception as e:
    print(f"An error occurred: {e}")
