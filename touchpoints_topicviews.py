import pandas as pd
import glob
import os

# Specify the folder paths
folder_paths = [
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - Jan',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - Feb',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - March',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - April',
    'C:\\Users\\Gajendra\\Desktop\\English_us\\TouchPoints - May'
]

# Initialize a dictionary to store the cumulative counts
usage_counts = {}

# Iterate over the folder paths
for month, folder_path in enumerate(folder_paths, start=1):
    # Find all Excel files in the current folder
    file_paths = glob.glob(folder_path + "/*.xlsx")
    
    # Iterate over the file paths
    for file_path in file_paths:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Extract the file name without extension as the row label
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # Calculate the total usage count for the current file
        total_count = df['Usage count'].sum()
        
        # Update the usage_counts dictionary
        if file_name in usage_counts:
            usage_counts[file_name][month] = total_count
        else:
            usage_counts[file_name] = {month: total_count}

# Create a DataFrame from the usage_counts dictionary
df_result = pd.DataFrame.from_dict(usage_counts, orient='index')

# Rearrange columns in the desired order (Jan, Feb, March, April, May)
column_order = [1, 2, 3, 4, 5]
df_result = df_result.reindex(columns=column_order)

# Rename the columns to the corresponding months
column_names = {
    1: "Jan",
    2: "Feb",
    3: "March",
    4: "April",
    5: "May"
}
df_result.rename(columns=column_names, inplace=True)

# Save the result to a new Excel file
df_result.to_excel("C:\\Users\\Gajendra\\Desktop\\usage_counts.xlsx", index_label="File Name")
