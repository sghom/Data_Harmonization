import pandas as pd

# Read the Excel sheets
spir_1 = pd.read_excel('Z:/ra/data_harmon/_match/spiromics_1.xlsx')
spir_comp = pd.read_excel('Z:/ra/data_harmon/spiromics/spiromics_comp_ext.xlsx')

# Extract rows corresponding to visit_1 in spir_1
visit_1_data = spir_1[spir_1['VISIT'] == 'VISIT_1']

# Identify columns in spir_comp that end with '_V1'
visit_1_columns = ['SUBJID'] + [col for col in spir_comp.columns if col.endswith('_V1')]

# Select these columns from spir_comp
spir_comp_visit_1 = spir_comp[visit_1_columns]

# Remove '_V1' from the column names
spir_comp_visit_1.columns = [col.replace('_V1', '') for col in spir_comp_visit_1.columns]

# Merge visit_1 data from spir_1 with spir_comp_visit_1 on SUBJID
merged_data = visit_1_data.merge(spir_comp_visit_1, on='SUBJID', how='left')

# Add the merged data back into the original spir_1 dataframe
# First, remove the original visit_1 rows from spir_1
spir_1_no_visit_1 = spir_1[spir_1['VISIT'] != 'VISIT_1']

# Concatenate the modified visit_1 data with the rest of spir_1
result_df = pd.concat([spir_1_no_visit_1, merged_data], ignore_index=True)

# Sort by SUBJID and VISIT to maintain the order
result_df = result_df.sort_values(by=['SUBJID', 'VISIT']).reset_index(drop=True)

# Replace duplicate SUBJID values with empty strings
result_df['SUBJID'] = result_df['SUBJID'].where(~result_df['SUBJID'].duplicated(), '')

# Save to an excel sheet
result_df.to_excel('Z:/ra/data_harmon/_match/spiromics_2.xlsx', index=False)

print(result_df)
