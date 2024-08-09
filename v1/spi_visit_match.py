### Match variables from V1, V2 and V3, identify matching variables in a separate columns.###
import numpy as np
import pandas as pd

# Specify the full paths to the Excel files
var_path = 'Z:/ra/data_harmon/var_iden.xlsx'

# Load the two Excel files
var_df = pd.read_excel(var_path, sheet_name='SPIROMICS', usecols=['CURRENT CORE VARIABLE NAME','LABEL'])
var_df = var_df.drop(columns=['LABEL']) #Dropping LABEL column. Remove this line, if you want to add correspoding variable definition for matched variables later on.

# Categorizing SPIROMICS varaibles by visit and creating new columns for each visit
vars = []
for i in range(1, 5):
    var_df_vi = var_df[var_df['CURRENT CORE VARIABLE NAME'].str.contains(f'_V{i}', na=False)].copy()
    var_df_vi.rename(columns={'CURRENT CORE VARIABLE NAME': f'spi_v{i}_var'}, inplace=True)
    vars.append(var_df_vi)
var_df = pd.concat([var_df] + vars, axis=1)


# Remove visit suffixes for matching
var_df['mod_spi_v1_var'] = var_df['spi_v1_var'].str.replace('_V1', '')
var_df['mod_spi_v2_var'] = var_df['spi_v2_var'].str.replace('_V2', '')
var_df['mod_spi_v3_var'] = var_df['spi_v3_var'].str.replace('_V3', '')
var_df['mod_spi_v4_var'] = var_df['spi_v4_var'].str.replace('_V4', '')

# Identify matched enteries between visits
match_v1_v2 = var_df['mod_spi_v1_var'].isin(var_df['mod_spi_v2_var'])
match_v1_v3 = var_df['mod_spi_v1_var'].isin(var_df['mod_spi_v3_var'])
match_v1_v4 = var_df['mod_spi_v1_var'].isin(var_df['mod_spi_v4_var'])

match_v2_v3 = var_df['mod_spi_v2_var'].isin(var_df['mod_spi_v3_var'])
match_v2_v4 = var_df['mod_spi_v2_var'].isin(var_df['mod_spi_v4_var'])

match_v3_v4 = var_df['mod_spi_v3_var'].isin(var_df['mod_spi_v4_var'])

# Create a new column to indicate matched entries
var_df['match_v1_v2'] = var_df['mod_spi_v1_var'].where(match_v1_v2, None)
var_df['match_v1_v3'] = var_df['mod_spi_v1_var'].where(match_v1_v3, None)
var_df['match_v1_v4'] = var_df['mod_spi_v1_var'].where(match_v1_v4, None)

var_df['match_v2_v3'] = var_df['mod_spi_v2_var'].where(match_v2_v3, None)
var_df['match_v2_v4'] = var_df['mod_spi_v2_var'].where(match_v2_v4, None)

var_df['match_v3_v4'] = var_df['mod_spi_v3_var'].where(match_v3_v4, None)

# Remove None (or NaN) values from match columns and reset the index
var_df['spi_v1_var'] = var_df['spi_v1_var'].dropna().reset_index(drop=True)
var_df['spi_v2_var'] = var_df['spi_v2_var'].dropna().reset_index(drop=True)
var_df['spi_v3_var'] = var_df['spi_v3_var'].dropna().reset_index(drop=True)
var_df['spi_v4_var'] = var_df['spi_v4_var'].dropna().reset_index(drop=True)
var_df['match_v1_v2'] = var_df['match_v1_v2'].dropna().reset_index(drop=True)
var_df['match_v1_v3'] = var_df['match_v1_v3'].dropna().reset_index(drop=True)
var_df['match_v1_v4'] = var_df['match_v1_v4'].dropna().reset_index(drop=True)
var_df['match_v2_v3'] = var_df['match_v2_v3'].dropna().reset_index(drop=True)
var_df['match_v2_v4'] = var_df['match_v2_v4'].dropna().reset_index(drop=True)
var_df['match_v3_v4'] = var_df['match_v3_v4'].dropna().reset_index(drop=True)

# Identifying the common variables between all visits
common = set(var_df['mod_spi_v1_var']).intersection(var_df['mod_spi_v2_var'], var_df['mod_spi_v3_var'], var_df['mod_spi_v4_var'])

# Convert all elements to strings first
common = [str(entry) for entry in common]

# Split values
common = [item for sublist in [entry.split(', ') for entry in common] for item in sublist] 

# Initialize new column with None values
var_df['match_all_visits'] = [None] * len(var_df)

# Fill new column with values from the list
var_df.loc[:len(common)-1, 'match_all_visits'] = common

# Replace 'nan' string with actual NaN
var_df['match_all_visits'] = var_df['match_all_visits'].replace('nan', np.nan)

# Remove None (or NaN) values from match columns and reset the index
var_df['match_all_visits'] = var_df['match_all_visits'].dropna().reset_index(drop=True)

# Drop the intermediate modified columns if needed
var_df.drop(['mod_spi_v1_var', 'mod_spi_v2_var','mod_spi_v3_var', 'mod_spi_v4_var'], axis=1, inplace=True)

print(var_df)

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/spi_visit_match.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    var_df.to_excel(writer, index=False, sheet_name='SPIROMICS')


