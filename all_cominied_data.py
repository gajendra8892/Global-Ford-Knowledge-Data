import pandas as pd
import os

# Set the path to the parent folder containing the language folders
parent_folder = "C:\\Users\\Gajendra\\Desktop\\New folder"

# Create an empty list to store all the dataframes
dfs = []

# Iterate over each language folder
for language_folder in os.listdir(parent_folder):
    language_folder_path = os.path.join(parent_folder, language_folder)
    
    # Iterate over each Excel file in the language folder
    for file in os.listdir(language_folder_path):
        if file.endswith(".xlsx"):  # Assuming all files are in XLSX format
            file_path = os.path.join(language_folder_path, file)
            
            # Read the Excel file and append it to the list of dataframes
            df = pd.read_excel(file_path)
            
            # Add an additional 'Touchpoint' column with the file name
            df['Touchpoint'] = os.path.splitext(file)[0]
            
            # Add an additional 'Locale' column with the folder name
            df['Locale'] = os.path.basename(language_folder_path)
            
            dfs.append(df)

# Concatenate all dataframes into a single dataframe
combined_df = pd.concat(dfs)

# Sort the combined dataframe by the 'topic_id' column
sorted_df = combined_df.sort_values('Topic Id')

# Split the sorted dataframe into multiple sheets based on the maximum allowed rows
max_rows_per_sheet = 1000000
num_sheets = (len(sorted_df) - 1) // max_rows_per_sheet + 1

# Create a new Excel file and save the data in multiple sheets
output_file = "C:\\Users\\Gajendra\\Desktop\\combined_data1.xlsx"
with pd.ExcelWriter(output_file) as writer:
    for i, sheet_num in enumerate(range(num_sheets)):
        start_row = i * max_rows_per_sheet
        end_row = (i + 1) * max_rows_per_sheet
        sheet_name = f'Sheet{i+1}'
        sorted_df.iloc[start_row:end_row].to_excel(writer, sheet_name=sheet_name, index=False)
