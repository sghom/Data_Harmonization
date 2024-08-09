### This script adds the pi10 and laa950 for the SPIROMICS dataset ###
### Also add the subject_id column to the SPIROMICS dataset sheet ###
import pandas as pd

# Read the first Excel sheet
df1 = pd.read_excel('Z:/ra/data_harmon/spiromics/spiromics_lla_pi10_cleaned.xlsx')

# Read the second Excel sheet
df2 = pd.read_excel('Z:/ra/data_harmon/spiromics/spiromics_missing.xlsx')

# List of visits to process
visits = ['VISIT_1', 'VISIT_4']

# List of columns to process from df1 and their corresponding suffixes for df2
columns_to_process = ['Pi10', 'Pi10(LEQ20)', 'LAA950']
suffixes = ['V1', 'V4']  # Suffixes for the columns in df2

# Iterate over each column and visit
for column in columns_to_process:
    for visit, suffix in zip(visits, suffixes):
        # Filter df1 to get rows where VISIT matches current visit
        df1_visit = df1[df1['VISIT'] == visit]
        
        # Merge df2 with the filtered df1 on the LAD4 column
        merged_df = df2.merge(df1_visit[['LAD4', column]], on='LAD4', how='left')
        
        # Update the corresponding column in df2 with the values from the merge
        df2[f'{column}_{suffix}'] = merged_df[column]

# Read the third Excel file (link between LAD4 an subjectids) containing subject_id and LAD4
file3_path = 'Z:/ra/data_harmon/spiromics/spiromics_link.xlsx'
df3 = pd.read_excel(file3_path)

# Merge based on the 'LAD4' column
merged_df_2 = pd.merge(df2, df3[['LAD4', 'subject_id']], on='LAD4', how='left')

#Reorder columns to have 'subject_id' as the first column
cols = merged_df_2.columns.tolist()
cols = ['subject_id'] + [col for col in cols if col != 'subject_id']
merged_df_2 = merged_df_2[cols]

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/_match/spiromics.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    merged_df_2.to_excel(writer, index=False, sheet_name='SPIROMICS')
