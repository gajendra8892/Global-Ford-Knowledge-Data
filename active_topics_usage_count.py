import pandas as pd
import glob
import os

# Get a list of all Excel files in the directory
excel_files = glob.glob('C:\\Users\\Gajendra\\Desktop\\TouchPoints - Jan\\*.xlsx')

# Create an empty DataFrame to store the combined data
combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Month', 'Usage count'])

# Iterate through each file
for file in excel_files:
    # Read the Excel file
    data = pd.read_excel(file)
    
    # Sort the data by Topic ID
    data_sorted = data.sort_values('Topic Id')
    
    # Extract the month from the file name (assuming the file name contains the month information)
    month = os.path.basename(file).split('.')[0]  # Extract only the month without the path
    
    # Add the month column to the sorted data
    data_sorted['Month'] = month
    
    # Update the combined data DataFrame
    combined_data = pd.concat([combined_data, data_sorted], ignore_index=True)

# Convert 'Usage Count' column to numeric values, ignoring non-numeric values
combined_data['Usage count'] = pd.to_numeric(combined_data['Usage count'], errors='coerce')

# Pivot the combined data to have separate columns for each month
pivot_table = combined_data.pivot_table(
    index=['Topic Id', 'Topic Name'],
    columns='Month',
    values='Usage count',
    aggfunc=lambda x: sum(x) if x.dtype != 'object' else '',
    fill_value=0
).reset_index()

# Calculate the total usage count across all months
pivot_table['Total Usage Count'] = pivot_table.drop(['Topic Id', 'Topic Name'], axis=1).sum(axis=1)

# Reorder the columns
columns_order = ['Topic Id', 'Topic Name'] + sorted(pivot_table.columns[2:-1]) + ['Total Usage Count']
pivot_table = pivot_table[columns_order]

# Filter the data to include only active topics with usage count >= 1
active_topics = pivot_table[pivot_table['Total Usage Count'] >= 1]

# Save the active topics to a new Excel file
active_topics.to_excel('C:\\Users\\Gajendra\\Desktop\\active_topics_usage_count.xlsx', index=False)
