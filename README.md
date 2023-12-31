# Python Excel Data Processing Project

## Introduction
This project assists a law firm in reading, processing, and formatting data from an Excel file containing decisions rendered by an international tribunal. It includes capabilities for data manipulation and chart creation.

## Prerequisites
- Python 3.x
- pandas
- openpyxl
- xlsxwriter
- Matplotlib (optional, if additional plotting functionalities are used)

## Installation
1. Clone or download the repository to your local machine.
2. Install the required packages:
   ```bash
   pip install pandas openpyxl xlsxwriter matplotlib
   ```

## Usage
To use the script, run `main.py` and follow the prompts to input the name of the source Excel file and the desired name for the output file.

```bash
python main.py
```

The script expects an Excel file with specific data formatting. Please ensure your input file meets the required structure.

## Project Structure
- `main.py`: The main script that orchestrates reading the input Excel file, processing the data, and writing to the output file.
- `utils.py`: Contains utility functions such as getting the executable path.
- `data_processing.py`: Includes the function `extract_decisions_and_cases` to process data within the Excel file.
- `excel_formatting.py`: Handles all formatting aspects for the Excel output, including cell styles, borders, alignments, and chart creation.

## Modules Description
- **utils**: 
  - `get_executable_path()`: Returns the absolute path of the running script.
- **data_processing**: 
  - `extract_decisions_and_cases(title)`: Extracts and processes decision and case information from a given title string in the data.
- **excel_formatting**: 
  - Contains functions to apply various formatting to an Excel worksheet, such as `apply_excel_formatting`, which adds borders, cell colors, adjusts column widths, and creates charts using `xlsxwriter`.

## Chart Creation
- This project now includes functionality to create charts in Excel files, leveraging the `xlsxwriter` library. This is particularly useful for visualizing data trends and creating histograms or other chart types directly in Excel reports.