# import os
# import pandas as pd

# # Function to get the usage count from a CAF, CAL, AAF, AAF1, or AAF2 Excel file
# def get_usage_count(file_path):
#     try:
#         df = pd.read_excel(file_path)
#         return df['Usage count'].sum()
#     except Exception as e:
#         return 0

# # Function to process a folder and its subfolders for CAF, CAL, AAF, AAF1, and AAF2 files
# def process_folder(folder_path):
#     caf_usage = 0
#     aaf_usage = 0
#     cal_usage = 0
#     for root, _, files in os.walk(folder_path):
#         for file in files:
#             if file.endswith(".xlsx"):
#                 file_path = os.path.join(root, file)
#                 if "CAF" in file:
#                     caf_usage += get_usage_count(file_path)
#                 elif "AAF1" in file or "SPOC" in file:
#                     aaf_usage += get_usage_count(file_path)
#                 elif "CAL" in file:
#                     cal_usage += get_usage_count(file_path)
#     return caf_usage, aaf_usage, cal_usage

# # Main function
# def main():
#     root_folder = "C:\\Users\\Gajendra\\Desktop\\Rest_of_America"
#     output_file = 'C:\\Users\\Gajendra\\Desktop\\Rest_of_America.xlsx'

#     data = []
#     caf_data = []
#     aaf_data = []
#     cal_data = []

#     month_order = {
#         "January": 1,
#         "February": 2,
#         "March": 3,
#         "April": 4,
#         "May": 5
#     }

#     for month_folder in sorted(os.listdir(root_folder), key=lambda x: month_order.get(x.split()[-1], 0)):
#         month_folder_path = os.path.join(root_folder, month_folder)
#         if os.path.isdir(month_folder_path):
#             caf_usage, aaf_usage, cal_usage = process_folder(month_folder_path)
#             month_name = month_folder.split()[-1]
#             data.append((month_name, caf_usage, aaf_usage, cal_usage))
#             caf_data.append((month_name, caf_usage))
#             aaf_data.append((month_name, aaf_usage))
#             cal_data.append((month_name, cal_usage))

#     # Save the data into an Excel file with separate sheets
#     with pd.ExcelWriter(output_file) as writer:
#         aaf_df = pd.DataFrame(aaf_data, columns=["Month", "AAF Usage count"])
#         aaf_df.to_excel(writer, sheet_name='AAF', index=False)
#         caf_df = pd.DataFrame(caf_data, columns=["Month", "CAF Usage count"])
#         caf_df.to_excel(writer, sheet_name='CAF', index=False)
#         cal_df = pd.DataFrame(cal_data, columns=["Month", "CAL Usage count"])
#         cal_df.to_excel(writer, sheet_name='CAL', index=False)

# if __name__ == "__main__":
#     main()



import os
import pandas as pd

# Function to get the usage count from a CAF, CAL, AAF, AAF1, or AAF2 Excel file
def get_usage_count(file_path):
    try:
        df = pd.read_excel(file_path)
        return df['Usage count'].sum()
    except Exception as e:
        return 0

# Function to process a folder and its subfolders for CAF, CAL, AAF, AAF1, and AAF2 files
def process_folder(folder_path, month):
    caf_usage = 0
    aaf_usage = 0
    cal_usage = 0
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                if "CAF" in file:
                    caf_usage += get_usage_count(file_path)
                elif "AAF1" in file or "SPOC" in file:
                    aaf_usage += get_usage_count(file_path)
                elif "CAL" in file:
                    cal_usage += get_usage_count(file_path)
    return month, aaf_usage, caf_usage, cal_usage

# Main function
def main():
    root_folder = "C:\\Users\\Gajendra\\Desktop\\Rest_of_America"
    output_file = 'C:\\Users\\Gajendra\\Desktop\\Rest_of_America.xlsx'

    data = []

    for month_folder in sorted(os.listdir(root_folder)):
        month_folder_path = os.path.join(root_folder, month_folder)
        if os.path.isdir(month_folder_path):
            month, aaf_usage, caf_usage, cal_usage = process_folder(month_folder_path, month_folder.split()[-1])
            data.append((month, aaf_usage, caf_usage, cal_usage))

    # Create a DataFrame and sort months in ascending order
    df = pd.DataFrame(data, columns=["Month", "AAF Usage count", "CAF Usage count", "CAL Usage count"])
    df['Month'] = pd.to_datetime(df['Month'], format='%B', errors='coerce')
    df.sort_values(by='Month', inplace=True, ignore_index=True)
    df['Month'] = df['Month'].dt.strftime('%B')

    # Save the data into an Excel file
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    main()



# import os
# import pandas as pd

# # Function to get the usage count from a CAF, CAL, AAF, AAF1, or AAF2 Excel file
# def get_usage_count(file_path):
#     try:
#         df = pd.read_excel(file_path)
#         return df['Usage count'].sum()
#     except Exception as e:
#         return 0

# # Function to process a folder and its subfolders for CAF, CAL, AAF, AAF1, and AAF2 files
# def process_subfolder(subfolder_path, month_folder_name, folder_name):
#     caf_usage = 0
#     aaf_usage = 0
#     cal_usage = 0
#     for root, _, files in os.walk(subfolder_path):
#         for file in files:
#             if file.endswith(".xlsx"):
#                 file_path = os.path.join(root, file)
#                 if "CAF" in file:
#                     caf_usage += get_usage_count(file_path)
#                 elif "AAF1" in file or "SPOC" in file:
#                     aaf_usage += get_usage_count(file_path)
#                 elif "CAL" in file:
#                     cal_usage += get_usage_count(file_path)
#     return month_folder_name, folder_name, aaf_usage, caf_usage, cal_usage

# # Main function
# def main():
#     root_folder = "C:\\Users\\Gajendra\\Desktop\\Rest_of_America"  # Replace with the actual path of your root folder
#     output_file = "C:\\Users\\Gajendra\\Desktop\\Rest_of_America_region_wise.xlsx"  # Name of the output Excel file

#     data = []

#     months = {
#         "January": 1,
#         "February": 2,
#         "March": 3,
#         "April": 4,
#         "May": 5,
#         "June": 6,
#         "July": 7
#     }

#     for month_folder in sorted(os.listdir(root_folder), key=lambda x: months[x.split()[0]]):
#         month_folder_path = os.path.join(root_folder, month_folder)
#         if os.path.isdir(month_folder_path):
#             for subfolder_name in sorted(os.listdir(month_folder_path)):
#                 subfolder_path = os.path.join(month_folder_path, subfolder_name)
#                 if os.path.isdir(subfolder_path):
#                     month_folder_name, folder_name, aaf_usage, caf_usage, cal_usage = process_subfolder(
#                         subfolder_path, month_folder, subfolder_name
#                     )
#                     data.append((month_folder_name, folder_name, aaf_usage, caf_usage, cal_usage))

#     # Create a DataFrame and sort months in ascending order
#     df = pd.DataFrame(data, columns=["Month", "Folder Name", "AAF Usage count", "CAF Usage count", "CAL Usage count"])
    
    
#     # Convert month names to datetime objects, sort, and convert back to strings
#     df['Month'] = pd.to_datetime(df['Month'], format='%B', errors='coerce')
#     df.sort_values(by='Month', inplace=True)
#     df['Month'] = df['Month'].dt.strftime('%B')

#     # Save the data into an Excel file with separate sheets for each subfolder
#     with pd.ExcelWriter(output_file) as writer:
#         for subfolder_name in df['Folder Name'].unique():
#             sub_df = df[df['Folder Name'] == subfolder_name].copy()
#             sub_df.drop(columns=['Folder Name'], inplace=True)
#             sub_df.to_excel(writer, sheet_name=subfolder_name.replace(" ", "_"), index=False)

# if __name__ == "__main__":
#     main()
