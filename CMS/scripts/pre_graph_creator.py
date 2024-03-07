import os
import pandas as pd
import matplotlib.pyplot as plt
from pandas.plotting import register_matplotlib_converters

register_matplotlib_converters()

# Update the file path
file_path = r'C:\Users\klongley\Documents\CMS_Testing\CMS\Pre_Site_Total_Alarms.xlsx'

# Read the Excel sheet
df = pd.read_excel(file_path, header=None)

# Convert the first column (column 0) to datetime format starting from row 1
df.iloc[1:, 0] = pd.to_datetime(df.iloc[1:, 0], format='%Y%m%d')

# Set the first row as column names
df.columns = df.iloc[0]

# Set the first column (date column) as the index
df.set_index(df.columns[0], inplace=True)

# Drop the first row after using it as column names
df = df.iloc[1:]

# Plot the graph with each site identified by column headers
ax = df.plot(kind='line', marker='o', figsize=(13, 6), legend=True)
ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%Y-%m-%d'))  # Format date on x-axis

# Save the figure in the specified folder
output_folder = r'C:\Users\klongley\Documents\CMS_Testing\CMS\graphs'
output_filename = 'Pre_Daily_Site_Alarms.png'
output_path = os.path.join(output_folder, output_filename)

# Check if the file exists, and replace if it does
if os.path.exists(output_path):
    os.remove(output_path)

plt.savefig(output_path)
plt.show()
