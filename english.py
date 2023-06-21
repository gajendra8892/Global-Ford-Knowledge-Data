import pandas as pd
import glob
import os


excel_files = glob.glob('C:\\Users\Gajendra\Desktop\English_us\\*.xlsx')


English_us_combined_data = pd.DataFrame(columns=['Topic Id', 'Topic Name', 'Month', 'Usage count'])

for file in excel_files:
    # Read the Excel file
    data = pd.read_excel(file)
    
    # Sort the data by Topic ID
    data_sorted = data.sort_values('Topic Id')
    
   
    month = os.path.basename(file).split('.')[0]  
    
    
    data_sorted['Month'] = month
    
    
    English_us_combined_data = pd.concat([English_us_combined_data, data_sorted], ignore_index=True)
    

pivot_table = English_us_combined_data.pivot_table(
    index=['Topic Id', 'Topic Name'],
    columns='Month',
    values='Usage count',
    aggfunc='sum',
    fill_value=0
).reset_index()


pivot_table.to_excel('C:\\Users\\Gajendra\\Desktop\\English_us_combined_data.xlsx', index=False)