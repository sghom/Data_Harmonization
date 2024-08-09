### Create a list of available scans for visits 1 and 3 in the CanCOLD Volumes folder ###

import os
import pandas as pd

# Define the folder path
path = 'R:/kirby_group/data/CanCOLD/ctVolumes'

# Get the folder names
folder_names = [f for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))]

# Initialize an empty list to hold the data
data = []

# Loop through each subject folder
for subject in folder_names:
    subject_path = os.path.join(path, subject)
    # Get subfolders that start with V1 or V3
    subfolders = [sf for sf in os.listdir(subject_path) if os.path.isdir(os.path.join(subject_path, sf)) and (sf.startswith('V1') or sf.startswith('V3'))]
    
    # Check if V1 and/or V3 exist and add to the data list
    for subfolder in subfolders:
        visit_number = subfolder[:2]  # Extract 'V1' or 'V3' from the subfolder name
        data.append([subject, visit_number])

# Create a DataFrame
img_df = pd.DataFrame(data, columns=['subjectId', 'VisitId'])

# Drop duplicate rows
img_df = img_df.drop_duplicates()

# Filter rows where visit number is 'V1' and 'V3'
img_df_v1 = img_df[img_df['VisitId'] == 'V1']
img_df_v3 = img_df[img_df['VisitId'] == 'V3']

# Save the DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/cancold/cancold_image_list.xlsx'

# Save the DataFrame to a new Excel file with the specified sheet name
with pd.ExcelWriter(output_path) as writer:
    img_df.to_excel(writer, sheet_name='Image_List', index=False)
    img_df_v1.to_excel(writer, sheet_name='Image_List_V1', index=False)
    img_df_v3.to_excel(writer, sheet_name='Image_List_V3', index=False)