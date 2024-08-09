### Match variables from V1, V2 and V3, identify matching variables in a separate columns.###

import pandas as pd

# Specify the full paths to the Excel files
v1_path = 'Z:/ra/data_harmon/can_v1_var.xlsx'
v2_path = 'Z:/ra/data_harmon/can_v2_var.xlsx'
v3_path = 'Z:/ra/data_harmon/can_v3_var.xlsx'
var_def_path = 'Z:/ra/data_harmon/var_iden.xlsx'

# Load the two Excel files
v1_df = pd.read_excel(v1_path, usecols=['can_v1_var'])
v2_df = pd.read_excel(v2_path, usecols=['can_v2_var'])
v3_df = pd.read_excel(v3_path, usecols=['can_v3_var'])

# Merge the two columns into a single DataFrame
merged_df = pd.concat([v1_df, v2_df, v3_df], axis=1)

# Convert all columns to lowercase for case-insensitive matching
merged_df = merged_df.apply(lambda x: x.astype(str).str.lower())

# Identify matched enteries between visits
match_v1_v2 = merged_df['can_v1_var'].isin(merged_df['can_v2_var'])
match_v1_v3 = merged_df['can_v1_var'].isin(merged_df['can_v3_var'])
match_v2_v3 = merged_df['can_v2_var'].isin(merged_df['can_v3_var'])


# Create a new column to indicate matched entries
merged_df['match_v1_v2'] = merged_df['can_v1_var'].where(match_v1_v2, None)
merged_df['match_v1_v3'] = merged_df['can_v1_var'].where(match_v1_v3, None)
merged_df['match_v2_v3'] = merged_df['can_v2_var'].where(match_v2_v3, None)

# Remove None (or NaN) values from match columns and reset the index
merged_df['match_v1_v2'] = merged_df['match_v1_v2'].dropna().reset_index(drop=True)
merged_df['match_v1_v3'] = merged_df['match_v1_v3'].dropna().reset_index(drop=True)
merged_df['match_v2_v3'] = merged_df['match_v2_v3'].dropna().reset_index(drop=True)

# Identifying the common variables between all visits
common = set(merged_df['can_v1_var']).intersection(merged_df['can_v2_var'], merged_df['can_v3_var'])
common = [item for sublist in [entry.split(', ') for entry in common] for item in sublist] # Converting set to list

# Initialize new column with None values
merged_df['match_all_visits'] = [None] * len(merged_df)

# Fill new column with values from the list
merged_df.loc[:len(common)-1, 'match_all_visits'] = common

# Getting rid of NaN values
merged_df = merged_df.fillna('')
merged_df = merged_df.replace('nan', '')

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/can_visit_match.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    merged_df.to_excel(writer, index=False, sheet_name='CanCOLD')

print(merged_df)
