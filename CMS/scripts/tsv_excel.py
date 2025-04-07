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

# Filter files with the correct date format and '.csv' extension
pre_filtered_files = [file for file in pre_files if file.endswith('.csv') and file.startswith('pre_alarms_')]
post_filtered_files = [file for file in post_files if file.endswith('.csv') and file.startswith('post_alarms_')]

# Print debugging information
print(f'All files in the pre directory: {pre_files}')
print(f'Filtered pre files: {pre_filtered_files}')
print(f'All files in the post directory: {post_files}')
print(f'Filtered post files: {post_filtered_files}')

# Function to process files and export to Excel
def process_file(file_path, excel_file_path):
    try:
        # Read the CSV file into a pandas DataFrame with specified encoding
        df = pd.read_csv(file_path, encoding='ISO-8859-1')  # or 'Windows-1252'

        # Write the DataFrame to an Excel file
        df.to_excel(excel_file_path, index=False)
        print(f'Data has been successfully imported from {file_path} to {excel_file_path}.')
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

# Process the most recent 'pre' file if available
if pre_filtered_files:
    try:
        most_recent_file = max(pre_filtered_files, key=lambda x: datetime.strptime(x[len('pre_alarms_'):-4], date_format))
        most_recent_file_path = os.path.join(pre_directory_path, most_recent_file)

        # Extract the date from the filename
        date_from_filename = most_recent_file[len('pre_alarms_'):-4]

        # Specify the Excel file path with the datestamped filename
        excel_file_path = f'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/pre/pre_alarms_{date_from_filename}.xlsx'

        # Process the file
        process_file(most_recent_file_path, excel_file_path)

    except ValueError as e:
        print(f"Error parsing date from the filename of the pre file: {e}")
else:
    print('No pre files matching the specified date format found.')

# Process the most recent 'post' file if available
if post_filtered_files:
    try:
        most_recent_file = max(post_filtered_files, key=lambda x: datetime.strptime(x[len('post_alarms_'):-4], date_format))
        most_recent_file_path = os.path.join(post_directory_path, most_recent_file)

        # Extract the date from the filename
        date_from_filename = most_recent_file[len('post_alarms_'):-4]

        # Specify the Excel file path with the datestamped filename
        excel_file_path = f'C:/Users/klongley/Documents/CMS_Testing/CMS/excel_exports/post/post_alarms_{date_from_filename}.xlsx'

        # Process the file
        process_file(most_recent_file_path, excel_file_path)

    except ValueError as e:
        print(f"Error parsing date from the filename of the post file: {e}")
else:
    print('No post files matching the specified date format found.')
