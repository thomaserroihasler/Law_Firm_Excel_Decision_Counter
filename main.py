import pandas as pd
from utils import get_executable_path
from data_processing import extract_decisions_and_cases
from excel_operations import remove_inner_borders, fill_cells, set_outer_border, write_in_cells, cell_alignment, auto_adjust_column_width
from collections import defaultdict

import xlsxwriter

def plot_frequency_histogram(workbook, worksheet_name, data, chart_title, start_cell='A1'):
    """
    Plots a frequency histogram in an Excel worksheet using xlsxwriter.

    Args:
        workbook (xlsxwriter.Workbook): An xlsxwriter Workbook object.
        worksheet_name (str): Name of the worksheet where the histogram will be plotted.
        data (list): A list of data points for the histogram.
        chart_title (str): Title of the histogram chart.
        start_cell (str, optional): The starting cell for plotting data. Defaults to 'A1'.
    """
    # Create a worksheet
    worksheet = workbook.add_worksheet(worksheet_name)

    # Write the data to the worksheet
    row = start_cell[1:]
    col = start_cell[0]
    for i, value in enumerate(data):
        worksheet.write(f'{col}{int(row) + i}', value)

    # Create a chart object
    chart = workbook.add_chart({'type': 'column'})

    # Configure the chart
    chart.add_series({
        'categories': f'={worksheet_name}!${col}${row}:${col}${int(row) + len(data) - 1}',
        'values': f'={worksheet_name}!${col}${row}:${col}${int(row) + len(data) - 1}',
        'gap': 2,
    })

    # Add a chart title
    chart.set_title({'name': chart_title})

    # Insert the chart into the worksheet
    worksheet.insert_chart(f'{col}{int(row) + len(data) + 2}', chart)

# Example usage
workbook = xlsxwriter.Workbook('histogram_example.xlsx')
data = [4, 5, 6, 7, 8, 9, 10, 11, 12]  # Example data
plot_frequency_histogram(workbook, 'Histogram', data, 'Sample Frequency Histogram')
workbook.close()



def main():
    print("Current Working Directory:", get_executable_path())


    input_file_path = input("Please enter the name of the excel file (without extension): ")
    output_file_path = input("Please enter the name of the saved excel file (without extension): ")


    input_file_path = input_file_path + '.xlsx'
    output_file_path = output_file_path + '.xlsx'
    try:
        df = pd.read_excel(get_executable_path() + '/' + input_file_path)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    df = df.iloc[:, [0, 2]]  # Selecting specific columns; adjust as needed
    df.iloc[:, 1] = pd.to_datetime(df.iloc[:, 1]).dt.date  # Convert to date-only format

    decision_case_counts = defaultdict(int)
    decisions = []

    for _, row in df.iterrows():
        decision_name = row.iloc[0]
        decision_date = row.iloc[1]
        cases = extract_decisions_and_cases(decision_name)
        cases_string = ' & '.join(cases)
        decisions.append([None, decision_name, decision_date, cases_string])  # Add None for the new first column
        decision_case_counts[(cases_string, decision_date)] += 1

    all_decisions_df = pd.DataFrame(decisions, columns=[None, 'Filename', 'Decision_Date', 'Cases_Numbers'])
    all_decisions_df = all_decisions_df.sort_values(by='Decision_Date', ascending=False)

    # Write to Excel
    with pd.ExcelWriter(get_executable_path() + '/' + output_file_path, engine='openpyxl') as writer:
        all_decisions_df.to_excel(writer, index=False, sheet_name='Cases_Numbers')

        workbook = writer.book
        worksheet = writer.sheets['Cases_Numbers']

        # Apply formatting, coloring, and border to decisions
        row_idx = 2  # Start from row 2 due to header
        counter = 1
        while row_idx <= len(all_decisions_df) + 1:
            row = all_decisions_df.iloc[row_idx - 2]
            decision_count = decision_case_counts[(row['Cases_Numbers'], row['Decision_Date'])]
            end_row = row_idx + decision_count - 1

            cell_range = f'B{row_idx}:D{end_row}'  # Adjust cell range
            cell_range_counter = f'A{row_idx}:A{end_row}'  # Adjust cell range
            remove_inner_borders(worksheet, cell_range)
            if end_row != row_idx:
                fill_cells(worksheet, cell_range, 'solid', '00FF00')
                set_outer_border(worksheet, cell_range, 'thick')


            write_in_cells(worksheet, cell_range_counter, str(counter))
            set_outer_border(worksheet, cell_range_counter, 'thick')

            row_idx = end_row + 1
            counter += 1
        cell_alignment(worksheet)
        auto_adjust_column_width(worksheet)
        set_outer_border(worksheet, f'B2:B{row_idx-1}', 'thin')
        set_outer_border(worksheet, f'C2:C{row_idx-1}', 'thin')
        set_outer_border(worksheet, None, 'thick')
        set_outer_border(worksheet, f'A1:D1', 'thick')

if __name__ == "__main__":
    main()