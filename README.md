# Instructions for Decision Counter

## Instructions

1. Ensure that the file to be treated is in the same folder as the program.
2. Open the program by clicking on the executable file.
3. Enter the name of the file to be treated (without the extension).
   - *Note: This file must exist; otherwise, the program will display a red error box.*
4. Enter the name of the file that will contain the decision counter (without the extension).
   - *Note: This file must not already exist; otherwise, the program will display a red error box.*
5. Click **Confirm**. The new file will be created in the folder as an Excel file.
6. Each case is grouped by color, such that if multiple cases exist, they are colored the same.
7. Archives and Bulls are extracted into two new worksheets.
8. In each worksheet, a bar chart displaying the number of unique decisions is provided.

## Input File Format

The input Excel file must have the following columns:
- **Filename**: This column contains the filenames and is used to identify and extract the case numbers.
- **Decision_Date**: This column contains the dates related to decisions.
- **Cases_Numbers**: This column contains the case numbers.

## Program Workflow

The program performs the following steps:
1. Reads the specified Excel file into a pandas DataFrame.
2. Extracts and writes case numbers based on patterns in the 'Filename' column.
3. Orders the DataFrame by 'Decision_Date' and groups similar cases together.
4. Groups identical decisions and creates a decision counter.
5. Writes the DataFrame to the output Excel file.
6. Separates the data into subcategories (Archives and Bulls) and writes them to respective worksheets.
7. Creates bar charts showing the number of unique case numbers per year for each worksheet.
8. Formats the worksheets to enhance readability and aesthetics.
