import pandas as pd
import glob
import os

# List of folders for different languages
folders = ["C:\\Users\\Gajendra\\Desktop\\New folder\\Arabic UAE", 
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Czech",
           "C:\\Users\Gajendra\\Desktop\\New folder\\Dannish",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Dutch Belgium",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Dutch Netherlands",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Argentina",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Australia",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Brazil",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Canada",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng India",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Ireland",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Mexico",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng NZ",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Philip",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng SA",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Thailand",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Turkey",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng UAE",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng UK",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng US",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Eng Vietnam",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Finnish",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\French Belgium",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\French Canada",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\French France",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\French Lux",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\French SWZL",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\German Austria",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\German Lux",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\German SWZL",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Germany German",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Greek Greece",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Hungarian",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Italian Italy",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Italian SWZL",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Norwegian",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Polish",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Portuguese Brazil",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Portuguese Portugal",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Romanian",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Slovenian",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Spanish Argentina",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Spanish Mexico",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Spanish Spain",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Spanish US",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Swedish",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Thai",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Turkish",
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Vietnamese Vietnam"]


# Create an empty DataFrame to store the combined data
combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'TouchPoint', 'Language', 'Usage count'])

# Iterate over the folders
for folder in folders:
    # Get the file paths for each folder
    file_paths = glob.glob(f'{folder}/*.xlsx')
    
    # Create a list to store the DataFrames for each file
    data_frames = []
    
    # Extract the language name from the folder
    language = os.path.basename(folder)
    
    # Iterate over the file paths
    for file_path in file_paths:
        # Extract the file name from the file path
        file_name = os.path.basename(file_path)
        touchpoint = os.path.splitext(file_name)[0].split(' - ')[-1]
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Filter the DataFrame for active touch points (Usage count > 0)
        active_touchpoints = df[df['Usage count'] > 0]
        
        # Add the month and language columns
        active_touchpoints['TouchPoint'] = touchpoint
        active_touchpoints['Language'] = language
        
        # Append the filtered DataFrame to the list
        data_frames.append(active_touchpoints)
    
    # Concatenate the DataFrames for the current folder vertically
    if data_frames:
        folder_data = pd.concat(data_frames, ignore_index=True)
        
        # Append the folder data to the combined_data DataFrame
        combined_data = pd.concat([combined_data, folder_data], ignore_index=True)

# Sort the combined_data DataFrame by Topic Id
combined_data.sort_values('Topic Id', inplace=True)

# Save the combined data to an Excel file
combined_data.to_excel('C:\\Users\\Gajendra\\Desktop\\combined_data2.xlsx', index=False)


