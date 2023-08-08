# import pandas as pd
# import glob
# import os

# # Get a list of all Excel files in the directory
# excel_files = glob.glob('C:\\Users\\Gajendra\\Desktop\\English_us\\English_US TouchPoints - Jan 2023\\*.xlsx')

# # Create an empty DataFrame to store the combined data
# combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Month', 'Usage count'])

# # Iterate through each file
# for file in excel_files:
#     # Read the Excel file
#     data = pd.read_excel(file)
    
#     # Sort the data by Topic ID
#     data_sorted = data.sort_values('Topic Id')
    
#     # Extract the month from the file name (assuming the file name contains the month information)
#     month = os.path.basename(file).split('.')[0]  # Extract only the month without the path
    
#     # Add the month column to the sorted data
#     data_sorted['Month'] = month
    
#     # Update the combined data DataFrame
#     combined_data = pd.concat([combined_data, data_sorted], ignore_index=True)
    
# # Pivot the combined data to have separate columns for each month
# pivot_table = combined_data.pivot_table(
#     index=['Topic Id', 'Topic Name'],
#     columns='Month',
#     values='Usage count',
#     aggfunc='sum',
#     fill_value=0
# ).reset_index()

# # Filter the pivot_table to include only active topics (usage count >= 1) for each month
# active_topics = pivot_table[(pivot_table.iloc[:, 2:] >= 1).any(axis=1)]

# # Save the active topics to a new Excel file
# active_topics.to_excel('C:\\Users\\Gajendra\\Desktop\\active_topics.xlsx', index=False)


import pandas as pd
import glob
import os
from openpyxl import Workbook

# Get a list of all folders containing Excel files
folders = ['C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - Jan',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - Feb',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - March',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - April',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - May']

# Create a new Excel workbook
workbook = Workbook()

# Iterate through each folder
for folder in folders:
    # Get a list of all Excel files in the folder
    excel_files = glob.glob(os.path.join(folder, '*.xlsx'))

    # Create an empty DataFrame to store the combined data
    combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Month', 'Usage count'])

    # Iterate through each file
    for file in excel_files:
        # Read the Excel file
        data = pd.read_excel(file)

        # Sort the data by Topic ID
        data_sorted = data.sort_values('Topic Id')

        # Extract the month from the file name
        month = os.path.basename(file).split('.')[0]

        # Add the month column to the sorted data
        data_sorted['Month'] = month

        # Update the combined data DataFrame
        combined_data = pd.concat([combined_data, data_sorted], ignore_index=True)

    # Pivot the combined data to have separate columns for each month
    pivot_table = combined_data.pivot_table(
        index=['Topic Id', 'Topic Name'],
        columns='Month',
        values='Usage count',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Filter the pivot_table to include only active topics (usage count >= 1) for each month
    active_topics = pivot_table[(pivot_table.iloc[:, 2:] >= 1).any(axis=1)]

    # Create a new sheet in the workbook
    sheet_name = os.path.basename(folder.rstrip('\\'))
    sheet = workbook.create_sheet(title=sheet_name)

    # Write the column names to the sheet
    sheet.append(active_topics.columns.tolist())

    # Write the active topics to the sheet
    for row in active_topics.values.tolist():
        sheet.append(row)

# Remove the default sheet created by openpyxl
workbook.remove(workbook['Sheet'])

# Save the workbook
workbook.save('C:\\Users\\Gajendra\\Desktop\\active_topics.xlsx')
