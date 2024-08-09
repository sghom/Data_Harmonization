### Match scans to variables for SPIROMICS ###
import pandas as pd

# Load data from Excel file
excel_file = 'Z:/ra/data_harmon/spiromics/spiromics_match1.xlsx'
df1 = pd.read_excel(excel_file, sheet_name='SPIROMICS')
df2 = pd.read_excel(excel_file, sheet_name='Image_Pull_List')

# Filter df1 based on matching entries
matched_df = df1[df1['subject_id'].isin(df2['matched_1'])]

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/_match/spiromics_matched.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    matched_df.to_excel(writer, index=False, sheet_name='SPIROMICS_MATCHED')