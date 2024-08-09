### Cleans up the variable dictionary provided by external datasets, allowing you to only focus on data variables provided and their definition###

import pandas as pd

# Paths to the existing Excel files containing variable definitions
cancold_xls = 'Z:\\ra\\data_harmon\\cancold\\CanCOLD_Variable_Def.xlsx'
spiromics_xls = 'Z:\\ra\\data_harmon\\spiromics\\SPIROMICS_Data_Dictionary.xlsx'

# Names of the sheets in the existing Excel files
cancold_sheet_name = "CRF V1V2 V3"
spiromics_sheet_name = "CORE6_4 Clinical Data Dictionar"

# Names of the columns to extract
cancold_columns_to_extract = ["Question", "Variables"]
spiromics_columns_to_extract = ["CURRENT CORE VARIABLE NAME", "LABEL"]

# Desired names for the new sheets in the new Excel file
cancold_new_sheet_name = "CanCOLD"
spiromics_new_sheet_name = "SPIROMICS"

# Desired path and name for the new Excel file
new_excel_path = "Z:\\ra\\data_harmon\\var_iden.xlsx"

# Read the existing Excel files, specifying the header row for CanCOLD
cancold_df = pd.read_excel(cancold_xls, sheet_name=cancold_sheet_name, header=1)
spiromics_df = pd.read_excel(spiromics_xls, sheet_name=spiromics_sheet_name)

# Extract the specified columns
cancold_extracted_columns_df = cancold_df[cancold_columns_to_extract]
spiromics_extracted_columns_df = spiromics_df[spiromics_columns_to_extract]

# Create a new Excel writer object
with pd.ExcelWriter(new_excel_path, engine='xlsxwriter') as writer:
    # Write the extracted columns to new sheets
    cancold_extracted_columns_df.to_excel(writer, sheet_name=cancold_new_sheet_name, index=False)
    spiromics_extracted_columns_df.to_excel(writer, sheet_name=spiromics_new_sheet_name, index=False)
