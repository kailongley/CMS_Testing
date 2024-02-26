import os
import pandas as pd

def generate_pre_site_total_alarms(pre_directory_path, output_file, location_order=None):
    # List all Excel files in the specified directory
    pre_excel_files = [file for file in os.listdir(pre_directory_path) if file.endswith('.xlsx')]

    # Initialize a dictionary to store the total alarms per site
    pre_total_alarms_dict = {}

    # Iterate through each Excel file in the directory
    for excel_file in pre_excel_files:
        pre_excel_file_path = os.path.join(pre_directory_path, excel_file)

        # Read the Excel file into a pandas DataFrame
        pre_new_data_df = pd.read_excel(pre_excel_file_path)

        # Extract the sheet name without the 'alarms_' prefix
        sheet_name = os.path.splitext(os.path.basename(excel_file))[0].replace('pre_alarms_', '')

        # Update the total alarms per site dictionary
        pre_total_alarms_dict[sheet_name] = pre_new_data_df.groupby('Location').size().to_dict()

    # If the dictionary is not empty, create a DataFrame
    if pre_total_alarms_dict:
        # Create a DataFrame using the dictionary
        pre_total_alarms_df = pd.DataFrame(pre_total_alarms_dict).T  # Transpose to have sheet names as rows

        # If location_order is provided, reorder the columns based on the specified order
        if location_order:
            pre_total_alarms_df = pre_total_alarms_df[location_order]

        # Fill missing values with 0
        pre_total_alarms_df = pre_total_alarms_df.fillna(0).astype(int)

        # Save the DataFrame to the output Excel file
        pre_total_alarms_df.to_excel(output_file, index=True, header=True)

        print(f'Total alarms per site has been successfully generated and saved in {output_file}.')
    else:
        print('No Excel files found in the directory or no "Location" column found. No data processed.')

def generate_post_site_total_alarms(post_directory_path, output_file, location_order=None):
    # List all Excel files in the specified directory
    post_excel_files = [file for file in os.listdir(post_directory_path) if file.endswith('.xlsx')]

    # Initialize a dictionary to store the total alarms per site
    post_total_alarms_dict = {}

    # Iterate through each Excel file in the directory
    for excel_file in post_excel_files:
        post_excel_file_path = os.path.join(post_directory_path, excel_file)

        # Read the Excel file into a pandas DataFrame
        post_new_data_df = pd.read_excel(post_excel_file_path)

        # Extract the sheet name without the 'alarms_' postfix
        sheet_name = os.path.splitext(os.path.basename(excel_file))[0].replace('post_alarms_', '')

        # Update the total alarms per site dictionary
        post_total_alarms_dict[sheet_name] = post_new_data_df.groupby('Location').size().to_dict()

    # If the dictionary is not empty, create a DataFrame
    if post_total_alarms_dict:
        # Create a DataFrame using the dictionary
        post_total_alarms_df = pd.DataFrame(post_total_alarms_dict).T  # Transpose to have sheet names as rows

        # If location_order is provided, reorder the columns based on the specified order
        if location_order:
            post_total_alarms_df = post_total_alarms_df[location_order]

        # Fill missing values with 0
        post_total_alarms_df = post_total_alarms_df.fillna(0).astype(int)

        # Save the DataFrame to the output Excel file
        post_total_alarms_df.to_excel(output_file, index=True, header=True)

        print(f'Total alarms per site has been successfully generated and saved in {output_file}.')
    else:
        print('No Excel files found in the directory or no "Location" column found. No data processed.')

if __name__ == "__main__":
    # Specify the path for the directory containing Excel files
    pre_excel_exports_directory = 'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/pre/'
    post_excel_exports_directory = 'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/post/'
    
    # Specify the output path for the new Excel files
    pre_output_excel_file = 'C:/Users/klongley/Documents/CMS_Testing/CMS/Pre_Site_Total_Alarms.xlsx'
    post_output_excel_file = 'C:/Users/klongley/Documents/CMS_Testing/CMS/Post_Site_Total_Alarms.xlsx'
    
    # Specify the desired order of Location names
    location_order = ['ARN', 'PKV', 'STL','THL','AZC','SDU']

    # Call the functions with the specified location order
    generate_pre_site_total_alarms(pre_excel_exports_directory, pre_output_excel_file, location_order)
    generate_post_site_total_alarms(post_excel_exports_directory, post_output_excel_file, location_order)


import openpyxl

def add_date_header(file_path):
    # Load the existing Excel file
    workbook = openpyxl.load_workbook(file_path)

    # Select the first sheet (assuming only one sheet is present)
    sheet = workbook.active

    # Add "Date" to cell A1
    sheet['A1'] = 'Date'

    # Save the changes
    workbook.save(file_path)

# Example usage:
add_date_header('C:/Users/klongley/Documents/CMS_Testing/CMS/Pre_Site_Total_Alarms.xlsx')
add_date_header('C:/Users/klongley/Documents/CMS_Testing/CMS/Post_Site_Total_Alarms.xlsx')
