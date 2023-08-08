import pandas as pd
import glob
import os

# List of folders for different languages
folders = ["C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Arabic UAE", 
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Czech",
           "C:\\Users\Gajendra\\Desktop\\Topic_usage_report_Feb\\Dannish",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Dutch Belgium",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Dutch Netherlands",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Argentina",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Australia",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Brazil",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Canada",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng India",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Ireland",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Mexico",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng NZ",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Philip",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng SA",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Thailand",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Turkey",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng UAE",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng UK",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng US",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Eng Vietnam",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Finnish",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\French Belgium",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\French Canada",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\French France",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\French Lux",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\French SWZL",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\German Austria",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\German Lux",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\German SWZL",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\German Germany",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Greek Greece",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Hungarian",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Italian Italy",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Italian SWZL",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Norwegian",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Polish",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Portuguese Brazil",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Portuguese Portugal",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Romanian",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Slovenian",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Spanish Argentina",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Spanish Mexico",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Spanish Spain",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Spanish US",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Swedish",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Thai",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Turkish",
           "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb\\Vietnamese Vietnam"]


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
combined_data.to_excel('C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb.xlsx', index=False)