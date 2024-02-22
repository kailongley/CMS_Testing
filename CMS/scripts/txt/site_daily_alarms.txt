import os
import pandas as pd
from datetime import datetime

def generate_site_total_alarms(directory_path, output_file):
    # List all Excel files in the specified directory
    excel_files = [file for file in os.listdir(directory_path) if file.endswith('.xlsx')]

    # Initialize a dictionary to store the total alarms per site
    total_alarms_dict = {}

    # Iterate through each Excel file in the directory
    for excel_file in excel_files:
        excel_file_path = os.path.join(directory_path, excel_file)

        # Read the Excel file into a pandas DataFrame
        new_data_df = pd.read_excel(excel_file_path)

        # Extract the sheet name without the 'alarms_' prefix
        sheet_name = os.path.splitext(os.path.basename(excel_file))[0].replace('alarms_', '')

        # Update the total alarms per site dictionary
        total_alarms_dict[sheet_name] = new_data_df.groupby('Location').size().to_dict()

    # If the dictionary is not empty, create a DataFrame
    if total_alarms_dict:
        # Create a DataFrame using the dictionary
        total_alarms_df = pd.DataFrame(total_alarms_dict).T  # Transpose to have sheet names as rows

        # Fill missing values with 0
        total_alarms_df = total_alarms_df.fillna(0).astype(int)

        # Save the DataFrame to the output Excel file
        total_alarms_df.to_excel(output_file, index=True, header=True)

        print(f'Total alarms per site has been successfully generated and saved in {output_file}.')
    else:
        print('No Excel files found in the directory or no "Location" column found. No data processed.')

if __name__ == "__main__":
    # Specify the path for the directory containing Excel files
    excel_exports_directory = 'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/'

    # Specify the output path for the new Excel file
    output_excel_file = 'C:/Users/klongley/Documents/CMS_Testing/CMS/Site_Total_Alarms.xlsx'

    # Call the function to generate the total alarms per site in a new Excel file
    generate_site_total_alarms(excel_exports_directory, output_excel_file)
