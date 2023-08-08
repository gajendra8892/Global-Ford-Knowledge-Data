# import pandas as pd
# import glob
# import os


# excel_files = glob.glob('C:\\Users\\Gajendra\\Desktop\\TouchPoints - Jan\\*.xlsx')

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

# # Sort the pivot_table by the sum of usage counts across all months in descending order
# pivot_table['Total Usage'] = pivot_table.iloc[:, 2:].sum(axis=1)
# pivot_table_sorted = pivot_table.sort_values('Total Usage', ascending=False)

# # Select the top 10 Topic IDs
# top_10_topics = pivot_table_sorted.head(10)

# # Save the top 10 topics to a new Excel file
# top_10_topics.to_excel('C:\\Users\\Gajendra\\Desktop\\top_10_topics.xlsx', index=False)


# import pandas as pd
# import glob
# import os

# excel_files = glob.glob('C:\\Users\\Gajendra\\Desktop\\TouchPoints - Jan\\*.xlsx')

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

# # Select only specific columns
# selected_columns = ['Topic Id', 'Topic Name', 'AAF1', 'CAF', 'CAL', 'DAF', 'SHFP', 'SHLW', 'SPOC']
# selected_topics = pivot_table[selected_columns]

# # Create a dictionary to store the top 10 topics for each column
# top_10_topics_all = {}

# # Get the top 10 topics based on the sum of usage counts for each column
# for column in selected_columns[2:]:
#     top_10_topics_col = selected_topics.nlargest(10, column)
#     top_10_topics_all[column] = top_10_topics_col

# # Save the top 10 topics to a single Excel file with separate sheets
# excel_file = 'C:\\Users\\Gajendra\\Desktop\\top_10_topics.xlsx'
# with pd.ExcelWriter(excel_file) as writer:
#     for column, topics in top_10_topics_all.items():
#         folder_name = os.path.basename(os.path.dirname(excel_files[0]))
#         sheet_name = f"{column.upper()}_{folder_name}"  # Use column name and folder name as sheet name
#         topics.to_excel(writer, sheet_name=sheet_name, index=False)

import pandas as pd
import glob
import os

folders = [
    "C:\\Users\\Gajendra\\Desktop\\New folder\\Arabic UAE", 
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
           "C:\\Users\\Gajendra\\Desktop\\New folder\\Vietnamese Vietnam"
    # Add more folder paths as needed
]

excel_file = 'C:\\Users\\Gajendra\\Desktop\\top_10_topics_combined.xlsx'
writer = pd.ExcelWriter(excel_file)

for folder in folders:
    excel_files = glob.glob(os.path.join(folder, '*.xlsx'))

    combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Month', 'Usage count'])

    for file in excel_files:
        # Read the Excel file
        data = pd.read_excel(file)

        # Sort the data by Topic ID
        data_sorted = data.sort_values('Topic Id')

        month = os.path.basename(file).split('.')[0]
        data_sorted['Month'] = month

        combined_data = pd.concat([combined_data, data_sorted], ignore_index=True)

    pivot_table = combined_data.pivot_table(
        index=['Topic Id', 'Topic Name'],
        columns='Month',
        values='Usage count',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Get the list of columns dynamically based on the files in the folder
    selected_columns = ['Topic Id', 'Topic Name'] + [os.path.splitext(os.path.basename(file))[0] for file in excel_files]

    # Select only the specific columns
    selected_topics = pivot_table[selected_columns]

    # Create a dictionary to store the top 10 topics for each column
    top_10_topics_all = {}

    # Get the top 10 topics based on the sum of usage counts for each column
    for column in selected_columns[2:]:
        top_10_topics_col = selected_topics.nlargest(10, column)
        top_10_topics_all[column] = top_10_topics_col

    for column, topics in top_10_topics_all.items():
        sheet_name = f"{column.upper()}_{os.path.basename(folder)}"
        topics.to_excel(writer, sheet_name=sheet_name, index=False)

writer._save()































