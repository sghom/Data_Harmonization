### This script prepares the pi10 and laa950 for the SPIROMICS dataset ###
### There are additional steps that need to be taken to pre-process the data ###

import pandas as pd

# Load the Excel file
file_path_1  = 'Z:/ra/data_harmon/spiromics/spiromics_lla_pi10.xlsx'  # Update with your file path
df = pd.read_excel(file_path_1 )

# Display the first few rows of the dataframe to understand its structure
print(df.head())

# Drop rows where any of the columns from the third to fifth are NaN
df_cleaned = df.dropna(subset=df.columns[2:5])

# Load the second Excel file (linking LAD4 ids to subjectids)
file_path_2 = 'Z:/ra/data_harmon/spiromics/spiromics_link.xlsx'  # Update with your file path
df_lad4 = pd.read_excel(file_path_2)

# Merge the dataframes on the 'subject_id' column
df_merged = pd.merge(df_cleaned, df_lad4[['subject_id', 'LAD4']], on='subject_id', how='left')

# Reorder columns to place LAD4 right after subject_id
cols = list(df_merged.columns)
subject_id_index = cols.index('subject_id')
cols.insert(subject_id_index + 1, cols.pop(cols.index('LAD4')))
df_merged = df_merged[cols]

# Display the cleaned dataframe
print(df_merged)

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/spiromics/spiromics_lla_pi10_cleaned.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    df_merged.to_excel(writer, index=False, sheet_name='SPIROMICS')