import pandas as pd
import numpy as np

# Read the Excel sheets
spir_df = pd.read_excel('Z:/ra/data_harmon/_match/spiromics_2.xlsx')
image_df = pd.read_excel('Z:/ra/data_harmon/spiromics/Image_Pull_List.xlsx')

# Identify SUBJID in spir_df that do not exist in image_df
subjid_not_in_image_df = spir_df[~spir_df['SUBJID'].isin(image_df['SUBJID'])]['SUBJID']
subjid_not_in_image_df = subjid_not_in_image_df.dropna()

# Temporarily Forward fill NaN values in SUBJID column
spir_df['SUBJID'] = spir_df['SUBJID'].fillna(method='ffill')

# Filter df1 to remove rows with SUBJID in subjid_to_remove
spir_df = spir_df[~spir_df['SUBJID'].isin(subjid_not_in_image_df)]

# Filter rows where visit number is 'visit_1'
spir_df_v1 = spir_df[spir_df['VISIT'] == 'VISIT_1']

# Replace duplicate SUBJID values with empty strings
spir_df_v1['SUBJID'] = spir_df_v1['SUBJID'].where(~spir_df_v1['SUBJID'].duplicated(), '')

# Remove rows where 'Column_Name' is ''
spir_df_v1 = spir_df_v1[spir_df_v1['SUBJID'] != '']

# Filter rows where visit number is 'visit_4'
spir_df_v4 = spir_df[spir_df['VISIT'] == 'VISIT_4']

# Replace duplicate SUBJID values with empty strings
spir_df['SUBJID'] = spir_df['SUBJID'].where(~spir_df['SUBJID'].duplicated(), '')

# Extract and store letters from the SUBJID to create Centre_ID
spir_df['Centre_ID'] = spir_df['SUBJID'].str.extract('([A-Za-z]+)', expand=False)
spir_df_v1['Centre_ID'] = spir_df_v1['SUBJID'].str.extract('([A-Za-z]+)', expand=False)
spir_df_v4['Centre_ID'] = spir_df_v4['SUBJID'].str.extract('([A-Za-z]+)', expand=False)

# Remove letters from the Subject ID column
spir_df['SUBJID'] = spir_df['SUBJID'].str.replace('[A-Za-z]', '', regex=True)
spir_df_v1['SUBJID'] = spir_df_v1['SUBJID'].str.replace('[A-Za-z]', '', regex=True)
spir_df_v4['SUBJID'] = spir_df_v4['SUBJID'].str.replace('[A-Za-z]', '', regex=True)

# Reorder columns to place Centre_ID after SUBJID
def reorder_columns(df):
    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index('Centre_ID')))
    return df[cols]

spir_df = reorder_columns(spir_df)
spir_df_v1 = reorder_columns(spir_df_v1)
spir_df_v4 = reorder_columns(spir_df_v4)

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/_match/spiromics_matched.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    spir_df.to_excel(writer, index=False, sheet_name='SPIROMICS_Matched')
    spir_df_v1.to_excel(writer, sheet_name='V1', index=False)
    spir_df_v4.to_excel(writer, sheet_name='V4', index=False)
