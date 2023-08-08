# import pandas as pd

# # Replace 'input_file.xlsx' with the name of your original Excel file
# input_file = "C:\\Users\\Gajendra\\Desktop\\US_English\\Topic_wise_combined_data.xlsx"
# output_file = "C:\\Users\\Gajendra\\Desktop\\top_20_topics.xlsx"

# # Read the original Excel file into a pandas DataFrame
# df = pd.read_excel(input_file)

# # Replace 'SYNC-Navigation Map Updates (Index)' with the desired topic name
# topic_name = 'SYNC - Navigation Map Updates (Index)'

# # Filter rows for the specified topic
# filtered_df = df[df['Topic Name'] == topic_name]

# # Group by 'Touchpoint' and 'Locale' and calculate the sum of 'Usage count' for each touchpoint and locale combination
# grouped_df = filtered_df.groupby(['Touchpoint', 'Locale'], as_index=False)['Usage count'].sum()

# # Calculate the total usage count for each touchpoint
# touchpoint_totals = grouped_df.groupby('Touchpoint')['Usage count'].sum().reset_index()

# # Calculate the grand total usage count for all touchpoints
# grand_total_usage = touchpoint_totals['Usage count'].sum()

# # Create a new DataFrame for the desired output format
# output_df = pd.DataFrame(columns=['Touchpoint', 'Locale', 'Usage count'])

# for touchpoint, group in grouped_df.groupby('Touchpoint'):
#     first_locale = True
#     for _, row in group.iterrows():
#         if first_locale:
#             output_df = output_df._append(row, ignore_index=True)
#             first_locale = False
#         else:
#             output_df = output_df._append({'Touchpoint': '', 'Locale': row['Locale'], 'Usage count': row['Usage count']}, ignore_index=True)

#     touchpoint_total = touchpoint_totals[touchpoint_totals['Touchpoint'] == touchpoint]['Usage count'].iloc[0]
#     output_df = output_df._append({'Touchpoint': f'{touchpoint} Total', 'Locale': '', 'Usage count': touchpoint_total}, ignore_index=True)

# # Add the "Grand Total" row
# output_df = output_df._append({'Touchpoint': 'Grand Total', 'Locale': '', 'Usage count': grand_total_usage}, ignore_index=True)

# # Save the result to a new Excel file
# output_df.to_excel(output_file, index=False)

# print("Excel file with touchpoints, locale, and usage count for the specified topic has been created.")

import pandas as pd
import re

# Function to replace invalid characters in a topic name with an underscore
def clean_topic_name(topic):
    return re.sub(r'[^\w\s]', '_', topic)[:31]

# Replace 'input_file.xlsx' with the name of your original Excel file
input_file = "C:\\Users\\Gajendra\\Desktop\\US_English\\Topic_wise_combined_data.xlsx"
output_file = "C:\\Users\\Gajendra\\Desktop\\top_20_topics.xlsx"

# Read the original Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

# List of 20 topics (replace these with your actual topic names)
topics = ['FPLW - Top 10 My Account',
          'FPLW - Activate FordPass/Lincoln Connect',
          'SYNC - Check for SYNC Software Updates',
          'MSI - Vehicle Order Status',
          'FPLW - FordPass/Lincoln Way App',
          'SYNC - Navigation Map Updates (Index)',
          'Vehicle - VIN Location',
          'Recall - Vehicle Involved In Recall',
          'FPLW - Remote Start Using the FordPass App',
          'SYNC - Performing a SYNC Master/Factory Reset (Index)',
          'FPLW - FordPass/Lincoln Way Connect Availability',
          'SYNC - Pairing a Phone with SYNC (Index)',
          'XX - SYNC - How To Get Automatic SYNC 3 WiFi Updates',
          'SYNC - Checking Your SYNC Software Version (Index)',
          'Ford Credit - Where can I find my account number?',
          'Ford Credit - Where do I send my payoff check?',
          'Recall - 22S73:Bronco Sport/Escape - Fuel Injector',
          'Ford Credit - Can I defer/extend a payment on my account?',
          'Feature - Using Remote Start (Index)',
          'Part - Retrieving the Keyless Entry Code (Smart Key)']

# Create an ExcelWriter object to write to the same Excel file
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    for topic in topics:
        # Clean the topic name to replace invalid characters and truncate to 31 characters
        sheet_name = clean_topic_name(topic)

        # Filter rows for the current topic
        filtered_df = df[df['Topic Name'] == topic]

        # Group by 'Touchpoint' and 'Locale' and calculate the sum of 'Usage count' for each touchpoint and locale combination
        grouped_df = filtered_df.groupby(['Touchpoint', 'Locale'], as_index=False)['Usage count'].sum()

        # Calculate the total usage count for each touchpoint
        touchpoint_totals = grouped_df.groupby('Touchpoint')['Usage count'].sum().reset_index()

        # Calculate the grand total usage count for all touchpoints
        grand_total_usage = touchpoint_totals['Usage count'].sum()

        # Create a new DataFrame for the desired output format
        output_df = pd.DataFrame(columns=['Touchpoint', 'Locale', 'Usage count'])

        for touchpoint, group in grouped_df.groupby('Touchpoint'):
            first_locale = True
            for _, row in group.iterrows():
                if first_locale:
                    output_df = output_df._append(row, ignore_index=True)
                    first_locale = False
                else:
                    output_df = output_df._append({'Touchpoint': '', 'Locale': row['Locale'], 'Usage count': row['Usage count']}, ignore_index=True)

            touchpoint_total = touchpoint_totals[touchpoint_totals['Touchpoint'] == touchpoint]['Usage count'].iloc[0]
            output_df = output_df._append({'Touchpoint': f'{touchpoint} Total', 'Locale': '', 'Usage count': touchpoint_total}, ignore_index=True)

        # Add the "Grand Total" row
        output_df = output_df._append({'Touchpoint': 'Grand Total', 'Locale': '', 'Usage count': grand_total_usage}, ignore_index=True)

        # Write the DataFrame to a new sheet with the cleaned topic name
        output_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)  # Start writing from the second row

        # Write the topic name as a header in the first row of the sheet
        worksheet = writer.sheets[sheet_name]
        worksheet.write(0, 0, topic)

print("Excel file with touchpoints, locale, and usage count for all topics has been created.")
