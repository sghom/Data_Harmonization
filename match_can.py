### Matching CanCOLD data with the available scans for each visit###

import pandas as pd
import numpy as np

# Path to your Excel file
can_path = 'Z:/ra/data_harmon/_match/cancold.xlsx'
can_img_list_path = 'Z:/ra/data_harmon/cancold/cancold_image_list.xlsx'

# Specify the sheet name you want to read
can_v1_sheet = 'CanCOLD_V1'  # Replace with the actual sheet name
can_v3_sheet = 'CanCOLD_V3'  # Replace with the actual sheet name

# Read the specified sheet into a DataFrame
can_df_v1 = pd.read_excel(can_path, sheet_name=can_v1_sheet)
can_df_v3 = pd.read_excel(can_path, sheet_name=can_v3_sheet)

# Specify the sheet name you want to read from Image list
img_list_v1 = 'Image_List_V1'  # Replace with the actual sheet name
img_list_v3 = 'Image_List_V3'  # Replace with the actual sheet name

# Read the specified sheet into a DataFrame
img_df_v1 = pd.read_excel(can_img_list_path, sheet_name=img_list_v1)
img_df_v3 = pd.read_excel(can_img_list_path, sheet_name=img_list_v3)

# Filter can_df_v1 based on subjectids in img_df_v1
can_df_v1 = can_df_v1[can_df_v1['subjectId'].isin(img_df_v1['subjectId'])]

# Filter can_df_v3 based on subjectids in img_df_v3
can_df_v3 = can_df_v3[can_df_v3['subjectId'].isin(img_df_v3['subjectId'])]

# Save the DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/_match/cancold_matched.xlsx'

# Save the DataFrame to a new Excel file with the specified sheet name
with pd.ExcelWriter(output_path) as writer:
    can_df_v1.to_excel(writer, sheet_name='CanCOLD_V1_Matched', index=False)
    can_df_v3.to_excel(writer, sheet_name='CanCOLD_V3_Matched', index=False)