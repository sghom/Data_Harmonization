# Data Harmonization
Loading Data from Various Sources and Organizing it in One Universal Excel File
This script performs the following tasks:

1. Load Data from Excel Sheets:
Reads multiple Excel files containing CanCOLD data and exacerbation records.
Filters and merges the data based on 'subjectId' to create a unified dataset.

2. Create a List of Available Scans:
Scans the ctVolumes folder to compile a list of available scans for visits 'V1' and 'V3'.
Creates and saves a DataFrame listing all subjects and their corresponding visit numbers.

3. Match CanCOLD Data with Available Scans:
Filters the CanCOLD data based on the availability of scans for each visit.
Saves the matched data into a new Excel file.

4. Process SPIROMICS Data:
Reads SPIROMICS data and an image pull list.
Filters out subjects not present in the image pull list and processes visit-specific data.
Updates columns based on visit-specific data and saves the processed data to an Excel file.

5. Merge Additional SPIROMICS Data:
Merges additional SPIROMICS data with previously processed data.
Updates columns based on visit-specific information and saves the final processed data.

Each section is carefully designed to handle specific data requirements and ensure that the datasets are organized and ready for further analysis.
