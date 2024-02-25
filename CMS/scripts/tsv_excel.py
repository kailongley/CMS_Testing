import os
import pandas as pd
from datetime import datetime

# Specify the directory where your files are located
pre_directory_path = 'C:/Users/klongley/Documents/CMS_Testing/CMS/raw_exports/pre/'
post_directory_path = 'C:/Users/klongley/Documents/CMS_Testing/CMS/raw_exports/post/'
# List all files in the directory
pre_files = os.listdir(pre_directory_path)
post_files = os.listdir(post_directory_path)
# Specify the date format in the filename
date_format = '%Y%m%d'  # Update this to match the actual date format

# Filter files with the correct date format and '.tsv' extension
pre_filtered_files = [file for file in pre_files if file.endswith('.tsv') and file.startswith('pre_alarms_')]
post_filtered_files = [file for file in post_files if file.endswith('.tsv') and file.startswith('post_alarms_')]
# Print debugging information
print(f'All files in the directory: {pre_files}')
print(f'Filtered files: {pre_filtered_files}')
print(f'All files in the directory: {post_files}')
print(f'Filtered files: {post_filtered_files}')
# Find the most recent file based on the date in the filename
if pre_filtered_files:
    most_recent_file = max(pre_filtered_files, key=lambda x: datetime.strptime(x[len('pre_alarms_'):-4], date_format))

    most_recent_file_path = os.path.join(pre_directory_path, most_recent_file)

    # Read the most recent TSV file into a pandas DataFrame
    df = pd.read_csv(most_recent_file_path, sep='\t')

    # Extract the date from the filename
    date_from_filename = most_recent_file[11:-4]

    # Specify the Excel file path with the datestamped filename
    excel_file_path = f'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/pre/pre_alarms_{date_from_filename}.xlsx'

    # Write the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

    print(f'Data has been successfully imported from {most_recent_file_path} to {excel_file_path}.')
else:
    print('No files matching the specified date format found in the directory.')
if post_filtered_files:
    most_recent_file = max(post_filtered_files, key=lambda x: datetime.strptime(x[len('post_alarms_'):-4], date_format))
    most_recent_file_path = os.path.join(post_directory_path, most_recent_file)

    # Read the most recent TSV file into a pandas DataFrame
    df = pd.read_csv(most_recent_file_path, sep='\t')

    # Extract the date from the filename
    date_from_filename = most_recent_file[11:-4]

    # Specify the Excel file path with the datestamped filename
    excel_file_path = f'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/post/post_alarms_{date_from_filename}.xlsx'

    # Write the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

    print(f'Data has been successfully imported from {most_recent_file_path} to {excel_file_path}.')
else:
    print('No files matching the specified date format found in the directory.')