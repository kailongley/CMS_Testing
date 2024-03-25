import os
import pandas as pd
import matplotlib.pyplot as plt
from pandas.plotting import register_matplotlib_converters

register_matplotlib_converters()

# Update the file path
file_path = r'C:\Users\klongley\Documents\CMS_Testing\CMS\Post_Site_Total_Alarms.xlsx'

# Read the Excel sheet
df = pd.read_excel(file_path)

# Convert the first column (column 0) to datetime format
df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format='%Y%m%d')

# Set the first column (date column) as the index
df.set_index(df.columns[0], inplace=True)

# Plot and save separate graphs for each site
output_folder = r'C:\Users\klongley\Documents\CMS_Testing\CMS\graphs\Sites'

for column in df.columns:
    # Create a new figure and axis for each site
    plt.figure(figsize=(12, 6))
    ax = plt.gca()
    
    # Plot data for the current site
    df[column].plot(ax=ax, kind='line', marker='o', legend=True)
    
    # Format date on x-axis
    ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%Y-%m-%d'))
    
    # Set plot title
    plt.title(f"Site: {column}")
    
    # Set plot labels
    plt.xlabel("Date")
    plt.ylabel("Alarms")
    
    # Save the figure with a unique filename for each site
    output_filename = f'{column}_Daily_Alarms.png'
    output_path = os.path.join(output_folder, output_filename)
    plt.savefig(output_path)
    
    # Close the plot to release memory
    plt.close()

print("Separate graphs for each site have been generated.")
