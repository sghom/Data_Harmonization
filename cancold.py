### Loading data from various sources and organizing it in one universal excel file ###

import pandas as pd

# Read the Excel sheets
can_v1 = pd.read_excel('Z:/ra/data_harmon/cancold/cancold_v1.xlsx')
can_v3 = pd.read_excel('Z:/ra/data_harmon/cancold/cancold_v3.xlsx')
can_exac = pd.read_excel('Z:/ra/data_harmon/cancold/CanCOLD_Exacerbations.xlsx')

# Visit 3 exacerbation data needs to be added using a script as the subjectids are not in the same order
# Filter rows where visit number is '3'
can_exac = can_exac[can_exac['VisitId'] == 3]

# Merging the dataframes on 'ID' and 'VISIT'
merged_df = pd.merge(can_v3, can_exac, on=['subjectId'], how='left')

# Drop the unnecessary 'VisitId' columns
merged_df.drop(columns='VisitId', inplace=True)

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/_match/cancold.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    can_v1.to_excel(writer, index=False, sheet_name='CanCOLD_V1')
    merged_df.to_excel(writer, index=False, sheet_name='CanCOLD_V3')
