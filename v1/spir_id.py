### FIND CURRENT SAMPLE IDs OF SPIROMICS DATASET ###
import os
import pandas as pd

# Define the path to the directory
directory_path = 'R:/kirby_group/data/SPRIOMICS/Volumes'

# Get the list of folder names in the directory
folder_names = [name for name in os.listdir(directory_path) if os.path.isdir(os.path.join(directory_path, name))]

# Create a DataFrame from the list of folder names
df = pd.DataFrame(folder_names, columns=['Subject_ID'])

# Define the path to the second Excel file
link_excel_path = 'Z:/ra/data_harmon/spiromics/linked.xlsx'

# Read the second Excel file into a DataFrame
link_df = pd.read_excel(link_excel_path)

# Extract the third column and the first column from the second DataFrame
third_column = link_df.iloc[:, 2]
first_column = link_df.iloc[:, 0]

# Initialize a list to store the matched values from the first column of the second DataFrame
matched_values = []

# Iterate over the folder names
for folder_name in df['Subject_ID']:
    # Find matches in the third column
    match = third_column[third_column.str.contains(folder_name, na=False)]
    
    if not match.empty:
        # Get the corresponding value from the first column (assuming the first match is taken if multiple)
        matched_value = first_column[match.index[0]]
    else:
        matched_value = None
    
    # Append the matched value to the list
    matched_values.append(matched_value)

# Add the matched values as a new column to the original DataFrame
df['SPIROMICS_subjid'] = matched_values

# Define the path to the Excel file
excel_file_path = 'Z:/ra/data_harmon/current_subject_ids.xlsx'

# Export the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)

# Display the DataFrame
print(df)
