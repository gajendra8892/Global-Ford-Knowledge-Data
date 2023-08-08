# import pandas as pd

# # Read the Excel file
# df = pd.read_excel("C:\\Users\\Gajendra\\Desktop\\topic_listing_with_usage_count.xlsx")

# # Print the loaded DataFrame
# print("Loaded DataFrame:")
# print(df)

# # Filter the data based on conditions
# filtered_df = df[df['Status'] == 'Pending Response']

# # Check if the filtered DataFrame is empty
# if filtered_df.empty:
#     print("No Pending Response topics found.")
# else:
#     # Convert 'Usage count' column to numeric
#     filtered_df['Usage count'] = pd.to_numeric(filtered_df['Usage count'], errors='coerce')

#     # Create a boolean mask for the comparison
#     viewed_mask = filtered_df['Usage count'] > 0

#     # Assign the values based on the boolean mask
#     filtered_df.loc[viewed_mask, 'Usage count'] = 'viewed'
#     filtered_df.loc[~viewed_mask, 'Usage count'] = 'not viewed'

#     # Calculate the sums of viewed and not viewed
#     sums = filtered_df.groupby('Usage count').size().reset_index(name='Sum')

#     # Concatenate the filtered DataFrame and sums as a new DataFrame
#     filtered_df = pd.concat([filtered_df, sums], ignore_index=True)

#     # Create a new Excel file with the filtered data and sums
#     filtered_df.to_excel('C:\\Users\\Gajendra\\Desktop\\pending_response_summary.xlsx', index=False)
#     print("Filtered data and sums have been written to output_file.xlsx.")

import pandas as pd

# Read the original Excel file
df = pd.read_excel("C:\\Users\\Gajendra\\Desktop\\pending_response_summary.xlsx")

# Convert the 'Touchpoint' column to string type
df['Touchpoint'] = df['Touchpoint'].astype(str)

# Create an empty DataFrame to store the results
new_df = pd.DataFrame(columns=['Touchpoint', 'viewed', 'not viewed'])

# Get unique touchpoints from the original DataFrame
touchpoints = df['Touchpoint'].unique()

# Iterate over each touchpoint
for touchpoint in touchpoints:
    if touchpoint != 'nan':
        # Filter the original DataFrame based on the current touchpoint
        touchpoint_df = df[df['Touchpoint'] == touchpoint]

        # Count the number of viewed and not viewed topics for the current touchpoint
        viewed_count = touchpoint_df[touchpoint_df['Usage count'] == 'viewed'].shape[0]
        not_viewed_count = touchpoint_df[touchpoint_df['Usage count'] == 'not viewed'].shape[0]

        # Append the touchpoint and counts to the new DataFrame
        new_df = new_df._append({'Touchpoint': touchpoint, 'viewed': viewed_count, 'not viewed': not_viewed_count},
                               ignore_index=True)

# Fill missing values with 0
new_df.fillna(0, inplace=True)

# Sort the DataFrame by touchpoint in alphabetical order
new_df.sort_values(by='Touchpoint', inplace=True)

# Calculate the sum of viewed and not viewed counts
sum_row = new_df[['viewed', 'not viewed']].sum().to_frame().T
sum_row['Touchpoint'] = 'sum'

# Concatenate the sum row with the rest of the DataFrame
new_df = pd.concat([new_df, sum_row], ignore_index=True)

# Save the new DataFrame to a new Excel file
new_df.to_excel("C:\\Users\\Gajendra\\Desktop\\pending_response.xlsx", index=False)