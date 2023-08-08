# import pandas as pd
# import os
# import glob

# # Path to the folder containing the Excel files
# folder_path = "C:\\Users\\Gajendra\\Desktop\\Topic_listing_Global_data"

# # List of columns to keep in the final output
# columns_to_keep = ['Topic ID', 'Topic Name', 'Touchpoint', 'Locale', 'Status']

# # Create an empty DataFrame to store the combined data
# combined_data = pd.DataFrame(columns=columns_to_keep)

# # Get the file paths of all Excel files in the folder
# file_paths = glob.glob(os.path.join(folder_path, "*.xlsx"))

# # Iterate over the file paths
# for file_path in file_paths:
#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(file_path)
    
#     # Filter the DataFrame for the desired columns
#     df = df[columns_to_keep]
    
#     # Append the DataFrame to the combined_data DataFrame
#     combined_data = pd.concat([combined_data, df], ignore_index=True)

# # Sort the combined_data DataFrame by Topic Id
# combined_data.sort_values('Topic ID', inplace=True)

# # Save the combined data to a new Excel file
# output_file_path = "C:\\Users\\Gajendra\\Desktop\\topic_listing_global_data.xlsx"
# combined_data.to_excel(output_file_path, index=False)

# import pandas as pd
# import os
# import glob

# # Path to the folder containing the Excel files
# folder_path = "C:\\Users\\Gajendra\\Desktop\\Topic_listing_Global_data"

# # List of columns to keep in the final output
# columns_to_keep = ['Topic ID', 'Topic Name', 'Touchpoint', 'Locale', 'Status']

# # Create an empty DataFrame to store the combined data
# combined_data = pd.DataFrame(columns=columns_to_keep)

# # Get the file paths of all Excel files in the folder
# file_paths = glob.glob(os.path.join(folder_path, "*.xlsx"))

# # Iterate over the file paths
# for file_path in file_paths:
#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(file_path)
    
#     # Filter the DataFrame for the desired columns
#     df = df[columns_to_keep]
    
#     # Append the filtered DataFrame to the combined_data DataFrame
#     combined_data = pd.concat([combined_data, df], ignore_index=True)

# # Sort the combined_data DataFrame by Topic ID
# combined_data.sort_values('Topic ID', inplace=True)

# # Create a dictionary to store the unique status for each topic ID
# topic_id_status = {}

# # Iterate over the combined_data DataFrame
# for index, row in combined_data.iterrows():
#     topic_id = row['Topic ID']
#     touchpoint = row['Touchpoint']
#     locale = row['Locale']
#     status = row['Status']
    
#     # If the topic ID is not in the dictionary, add it with the current status
#     if topic_id not in topic_id_status:
#         topic_id_status[topic_id] = (status, touchpoint, locale)
    
#     # If the topic ID is already in the dictionary, update the status only if it is not empty
#     elif topic_id_status[topic_id][0] != status and pd.notna(status):
#         topic_id_status[topic_id] = (status, touchpoint, locale)

# # Update the status column in the combined_data DataFrame
# combined_data['Status'] = combined_data.apply(lambda row: '' if (row['Touchpoint'], row['Locale']) != topic_id_status[row['Topic ID']][1:] else topic_id_status[row['Topic ID']][0], axis=1)

# # Save the combined_data DataFrame to a new Excel file
# output_file_path = "C:\\Users\\Gajendra\\Desktop\\topic_listing_global_data3.xlsx"
# combined_data.to_excel(output_file_path, index=False)

import pandas as pd

# Read the Excel file
df = pd.read_excel("C:\\Users\\Gajendra\\Desktop\\topic_listing_with_usage_count.xlsx")

# Update 'Usage count' column based on the condition
df['Usage count'] = df['Usage count'].apply(lambda x: 'viewed' if x > 0 else 'not viewed')

# Save the modified dataframe to a new Excel file
df.to_excel("C:\\Users\\Gajendra\\Desktop\\topic_listing_with_usage_count1.xlsx", index=False)
