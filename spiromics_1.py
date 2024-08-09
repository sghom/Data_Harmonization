### Loading data from various sources and organizing it in one universal excel file ###

import pandas as pd

# Read the Excel sheets
spir_temp = pd.read_excel('Z:/ra/data_harmon/_match/spiromics_temp.xlsx')

spir_comp_laa_pi = pd.read_excel('Z:/ra/data_harmon/spiromics/spiromics_comp(laa_pi).xlsx')

spir_comp_gen = pd.read_excel('Z:/ra/data_harmon/spiromics/spiromics_comp(gen_rac_pex).xlsx')

# Remove empty duplicate rows
spir_comp_laa_pi = spir_comp_laa_pi.dropna()

# Merging the dataframes on 'ID' and 'VISIT'
merged_df = pd.merge(spir_temp, spir_comp_laa_pi, on=['SUBJID', 'VISIT'], how='left')

merged_df['LAA950'] = merged_df['BOTH_PCT_BELOW_EQUAL_950']
merged_df['Pi10'] = merged_df['WHOLE_TREE_ALL']
merged_df['Pi10(leq20)'] = merged_df['WHOLE_TREE_LEQ20']

# Drop the unnecessary merged columns
merged_df.drop(columns=['BOTH_PCT_BELOW_EQUAL_950', 'WHOLE_TREE_ALL','WHOLE_TREE_LEQ20'], inplace=True)

# Merge the dataframes on 'subject_id' adding race and geder columns
merged_df_2 = pd.merge(merged_df, spir_comp_gen, on='SUBJID', how='left')

merged_df_2['GENDER'] = merged_df_2['gender']
merged_df_2['RACE'] = merged_df_2['race']

# Drop the unnecessary merged columns
merged_df_2.drop(columns=['gender', 'race'], inplace=True)

# # Replace duplicate SUBJID values with empty strings
# merged_df_2['SUBJID'] = merged_df_2['SUBJID'].where(~merged_df_2['SUBJID'].duplicated(), '')

# List of columns to be updated based on duplicate SUBJID
columns_to_update = ['PEX_TOT_A','PEX_SEVERE_A','NO_ACUTE_EXAC','STRATUM_ENROLLED','STRATUM_GOLD','Y2_EXAC']

# Replace values in the specified columns with empty strings where SUBJID is empty
for col in columns_to_update:
    merged_df_2[col] = merged_df_2[col].where(merged_df_2['SUBJID'] != '', '')

# List of columns to be updated based on only variables that pretain to visit 1
columns_to_update_2 = ['PEX_TOT_V1','PEX_SEVERETOT_V1','PEX_DRUGTOT_V1','PEX_DRUG_NOHOS_V1','PEX_HOSP_V1']

# Replace values in the specified columns 'columns_to_update_2'
for col_2 in columns_to_update_2:
    merged_df_2[col_2] = merged_df_2[col_2].where(merged_df_2['VISIT'] == 'VISIT_1', '')

# List of columns to be updated based on only variables that pretain to visit 4
columns_to_update_3 = ['PEX_TOT_0_1095_A','PEX_SEVERE_0_1095_A','PEX_DRUG_0_1095_A','PEX_DRUG_NOHOS_0_1095_A','PEX_HOSP_0_1095_A']

# Replace values in the specified columns 'columns_to_update_2'
for col_3 in columns_to_update_3:
    merged_df_2[col_3] = merged_df_2[col_3].where(merged_df_2['VISIT'] == 'VISIT_4', '')

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/_match/spiromics_1.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    merged_df_2.to_excel(writer, index=False, sheet_name='SPIROMICS')