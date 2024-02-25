
import os
import pandas as pd

def append_to_master(master_excel_file, new_excel_file):
    # Initialize a DataFrame for the master file or create an empty DataFrame if the file doesn't exist
    if os.path.exists(master_excel_file):
        master_df = pd.read_excel(master_excel_file)
    else:
        master_df = pd.DataFrame()

    # Read the new Excel file into a pandas DataFrame
    new_data_df = pd.read_excel(new_excel_file)

    # Identify new rows in the new_data_df
    new_rows = new_data_df[~new_data_df.isin(master_df.to_dict(orient='list')).all(axis=1)]

    # Check if there are new rows
    if not new_rows.empty:
        # Append the new rows to the master DataFrame
        master_df = pd.concat([master_df, new_rows], ignore_index=True)

        # Update the master Excel file
        master_df.to_excel(master_excel_file, index=False)

        print(f'Data has been successfully appended from {new_excel_file} to {master_excel_file}.')
    else:
        print('No new data found in the new file. No append performed.')

if __name__ == "__main__":
    # Specify the path for the master Excel file
    master_excel_file_path = 'C:/Users/klongley/Documents/CMS_Testing/CMS/Master_Alarms.xlsx'
    
    # Find the most recent Excel file based on the alphanumeric order of filenames
    directory_path = 'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/pre/'
    excel_files = [file for file in os.listdir(directory_path) if file.startswith('pre_alarms_') and file.endswith('.xlsx')]
    
    if excel_files:
        most_recent_file = max(excel_files, key=lambda x: x[7:-5])
        most_recent_file_path = os.path.join(directory_path, most_recent_file)

        # Call the function to append data to the master Excel file
        append_to_master(master_excel_file_path, most_recent_file_path)
    else:
        print('No files matching the specified format found in the directory.')