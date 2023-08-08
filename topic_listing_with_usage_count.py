import pandas as pd

# Read the first Excel file
df1 = pd.read_excel("C:\\Users\\Gajendra\\Desktop\\topic_listing_global_data.xlsx")

# Read the second Excel file
df2 = pd.read_excel("C:\\Users\\Gajendra\\Desktop\\US_English\\Topic_wise_combined_data.xlsx")

# Merge the two dataframes on 'topic id', 'topic name', 'touchpoint', and 'locale'
merged_df = pd.merge(df1, df2, on=['Topic Id', 'Topic Name', 'Touchpoint', 'Locale'], how='left')

# Fill missing usage count with 0
merged_df['Usage count'] = merged_df['Usage count'].fillna(0)

# Save the merged dataframe to a new Excel file
merged_df.to_excel('C:\\Users\\Gajendra\\Desktop\\topic_listing_with_usage_count.xlsx', index=False)

# import pandas as pd

# # File paths
# file_paths = [
#     "C:\\Users\\Gajendra\\Desktop\\topic_listing_global_data.xlsx",
#     "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Jan.xlsx",
#     "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_Feb.xlsx", 
#     "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_March.xlsx", 
#     "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_April.xlsx", 
#     "C:\\Users\\Gajendra\\Desktop\\Topic_usage_report_May.xlsx"
# ]

# # Read the first Excel file
# df_main = pd.read_excel(file_paths[0])

# # Iterate over the remaining files
# for path in file_paths[1:]:
#     # Extract the file name without extension
#     file_name = path.split('.')[0]
#     # Extract the last word of the file name
#     column_name = file_name.split('_')[-1]
    
#     # Read the file
#     df_temp = pd.read_excel(path)
    
#     # Merge the data based on 'Topic Id', 'Touchpoint', and 'Locale'
#     df_main = pd.merge(df_main, df_temp[['Topic Id', 'Touchpoint', 'Locale', 'Usage count']],
#                        on=['Topic Id', 'Touchpoint', 'Locale'], how='left')
    
#     # Rename the 'Usage count' column with the file name
#     df_main = df_main.rename(columns={'Usage count': column_name})
    
#     # Fill missing values with zero
#     df_main[column_name].fillna(0, inplace=True)

# # Save the merged data to a new Excel file
# df_main.to_excel("C:\\Users\\Gajendra\\Desktop\\Topic_listing_usage_count.xlsx", index=False)


