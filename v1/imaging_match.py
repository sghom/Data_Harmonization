### Match the imaging data avaialble with SPIROMICS data available ###
import pandas as pd
import os

# Read the SPIROMICS Excel file
file1_path = 'Z:/ra/data_harmon/_match/spiromics.xlsx'
df1 = pd.read_excel(file1_path)

# Read the Image Pull List Excel file
file2_path = 'Z:/ra/data_harmon/spiromics/Image_Pull_List.xlsx'
df2 = pd.read_excel(file2_path)

# Rename columns
df2.rename(columns={'subject_id': 'image_pull_list_id'}, inplace=True)
df2.rename(columns={'subjid': 'subject_id'}, inplace=True)

# Remove columns
df2.drop(columns=['visit','session_id'], inplace=True)

# Extract the numeric part after 'H-'
df2['image_pull_list_id'] = df2['image_pull_list_id'].str.extract(r'H-(\d+)_?')

# Sort df2 by 'image_pull_list_id' and expand selection to other columns
df2 = df2.sort_values(by='image_pull_list_id').reset_index(drop=True)

# Directory path where your imaging data are located
directory_path = 'R:/kirby_group/data/SPRIOMICS/Scans/2_decomp_scans'

# Get all filenames in the directory
files = os.listdir(directory_path)

# Create a pandas DataFrame
img_df = pd.DataFrame(files, columns=['filename(decomp_folder)'])

# Concatenate the existing df2 with img_df
df2 = pd.concat([df2, img_df], axis=1)

# Create a dictionary mapping values from image_pull_list_id to subject_id
mapping_dict = dict(zip(df2['image_pull_list_id'], df2['subject_id']))

# Map values from filename(decomp_folder) using the dictionary and create a new column
df2['matched_1'] = df2['filename(decomp_folder)'].map(mapping_dict)

# Remove duplicates in 'column_name'
df2 = df2.drop_duplicates(subset=['matched_1'])

# Save the merged DataFrame to a new Excel file
output_path = 'Z:/ra/data_harmon/spiromics/spiromics_match1.xlsx'

# Create a Pandas Excel writer using xlsxwriter as the engine
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Write merged_df to a specific sheet
    df1.to_excel(writer, index=False, sheet_name='SPIROMICS')
    df2.to_excel(writer, index=False, sheet_name='Image_Pull_List')