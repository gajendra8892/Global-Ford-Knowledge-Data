# import pandas as pd
# import glob
# import os


# excel_files = glob.glob('C:\\Users\\Gajendra\\Desktop\\non english\\Non English-US Topic Usage - Jan 2023\\*.xlsx')


# English_us_combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Month', 'Usage count'])

# for file in excel_files:
#     # Read the Excel file
#     data = pd.read_excel(file)
    
#     # Sort the data by Topic ID
#     data_sorted = data.sort_values('Topic Id')
    
   
#     month = os.path.basename(file).split('.')[0]  
    
    
#     data_sorted['Month'] = month
    
    
#     English_us_combined_data = pd.concat([English_us_combined_data, data_sorted], ignore_index=True)
    

# pivot_table = English_us_combined_data.pivot_table(
#     index=['Topic Id', 'Topic Name'],
#     columns='Month',
#     values='Usage count',
#     aggfunc='sum',
#     fill_value=0
# ).reset_index()


# pivot_table.to_excel('C:\\Users\\Gajendra\\Desktop\\English_us_combined_data.xlsx', index=False)



import pandas as pd
import glob
import os
import xlsxwriter


# List of folders for different months
folders = [
     "C:\\Users\\Gajendra\\Desktop\\non english\\Non English US-Jan",
    "C:\\Users\\Gajendra\\Desktop\\non english\\Non English US-Feb",
    "C:\\Users\\Gajendra\\Desktop\\non english\\Non English US-March",
    "C:\\Users\\Gajendra\\Desktop\\non english\\Non English US-April",
    "C:\\Users\\Gajendra\\Desktop\\non english\\Non English US-May"
    # Add more folders for additional months here
]

# Create a new Excel file
writer = pd.ExcelWriter('C:\\Users\\Gajendra\\Desktop\\combined_data_non_english.xlsx', engine='xlsxwriter')

# Iterate over the folders
for folder in folders:
    # Extract the sheet name from the folder
    sheet_name = os.path.basename(folder)[:31]  # Limit the sheet name to 31 characters
    
    # Get the file paths for each folder
    excel_files = glob.glob(f'{folder}/*.xlsx')
    
    # Create an empty DataFrame to store the combined data for the current month
    combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Language', 'Usage count'])
    
    # Iterate over the Excel files
    for file in excel_files:
        # Read the Excel file into a DataFrame
        data = pd.read_excel(file)
        
        # Extract the language from the file name
        language = os.path.basename(file).split('.')[0]
        
        # Add the language column to the data
        data['Language'] = language
        
        # Append the data to the combined_data DataFrame
        combined_data = pd.concat([combined_data, data], ignore_index=True)
    
    # Pivot the combined_data DataFrame
    pivot_table = combined_data.pivot_table(
        index=['Topic Id', 'Topic Name'],
        columns='Language',
        values='Usage count',
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    # Write the pivot_table to a new sheet in the Excel file
    pivot_table.to_excel(writer, sheet_name=sheet_name, index=False)

# Close the Excel writer
writer.close()

pivot_table.to_excel('C:\\Users\\Gajendra\\Desktop\\English_us_combined_data.xlsx', index=False)