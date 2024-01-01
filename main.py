import pandas as pd
from collections import defaultdict
from utils import get_executable_path
from data_processing import extract_decisions_and_cases
from excel_operations import remove_inner_borders, fill_cells, set_outer_border
from excel_operations import write_in_cells, cell_alignment, auto_adjust_column_width
from excel_operations import     plot_frequency_bar_chart
from collections import Counter

def main():
    print("Current Working Directory:", get_executable_path())

    input_file_path = input("Please enter the name of the excel file (without extension): ")
    output_file_path = input("Please enter the name of the saved excel file (without extension): ")

    input_file_path += '.xlsx'
    output_file_path += '.xlsx'

    try:
        df = pd.read_excel(get_executable_path() + '/' + input_file_path)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    df = df.iloc[:, [0, 2]]  # Selecting specific columns
    df.iloc[:, 1] = pd.to_datetime(df.iloc[:, 1]).dt.date  # Convert to date-only format

    decision_case_counts = defaultdict(int)
    decisions = []

    for _, row in df.iterrows():
        decision_name = row.iloc[0]
        decision_date = row.iloc[1]
        cases = extract_decisions_and_cases(decision_name)
        cases_string = ' & '.join(cases)
        decisions.append([None, decision_name, decision_date, cases_string])
        decision_case_counts[(cases_string, decision_date)] += 1

    all_decisions_df = pd.DataFrame(decisions, columns=[None, 'Filename', 'Decision_Date', 'Cases_Numbers'])
    all_decisions_df = all_decisions_df.sort_values(by='Decision_Date', ascending=False)

    output_path = get_executable_path() + '/' + output_file_path
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
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

            cell_range = f'B{row_idx}:D{end_row}'
            cell_range_counter = f'A{row_idx}:A{end_row}'
            remove_inner_borders(worksheet, cell_range)
            if end_row != row_idx:
                fill_cells(worksheet, cell_range, 'solid', '00FF00')
                set_outer_border(worksheet, cell_range, 'thick')

            write_in_cells(worksheet, cell_range_counter, str(counter))
            set_outer_border(worksheet, cell_range_counter, 'thick')

            row_idx = end_row + 1
            counter += 1

        cell_alignment(worksheet,'A')
        cell_alignment(worksheet, 'C')
        cell_alignment(worksheet, 'D')

        auto_adjust_column_width(worksheet)
        set_outer_border(worksheet, f'B2:B{row_idx-1}', 'thin')
        set_outer_border(worksheet, f'C2:C{row_idx-1}', 'thin')
        set_outer_border(worksheet, None, 'thick')
        set_outer_border(worksheet, f'A1:D1', 'thick')
        # Extract year from each datetime object
        decision_years = [date.year for date in all_decisions_df['Decision_Date'] if pd.notnull(date)]
        year_counts = Counter(decision_years)
        data_for_chart = list(year_counts.items())
        plot_frequency_bar_chart(workbook, 'Cases_Numbers', data_for_chart, 'Decision Date Frequency', 'F1')

        # Define subsets
        subset_1 = all_decisions_df[all_decisions_df['Filename'].str.contains('\[CAS Web Archives\]')]
        subset_2 = all_decisions_df[all_decisions_df['Filename'].str.contains('CAS Bull')]
        subset_3 = all_decisions_df[~all_decisions_df['Filename'].isin(subset_1['Filename']) & ~all_decisions_df['Filename'].isin(subset_2['Filename'])]

        # Write subsets to different sheets
        subsets = [subset_1,subset_2,subset_3]
        sheet_names = ['CAS Web Archives','CAS Bull','Other']
        for i, subset in enumerate(subsets, start=1):
            sheet_name = sheet_names[i-1]
            workbook.create_sheet(sheet_name)
            subset.to_excel(writer, index=False, sheet_name=sheet_name)

            subset_worksheet = writer.sheets[sheet_name]
            # Apply formatting and add counter to the first column
            row_idx = 2  # Start from the second row
            for _, row in subset.iterrows():
                write_in_cells(subset_worksheet, f'A{row_idx}', str(row_idx - 1))
                row_idx += 1
            cell_alignment(subset_worksheet, 'A')
            cell_alignment(subset_worksheet, 'C')
            cell_alignment(subset_worksheet, 'D')
            auto_adjust_column_width(subset_worksheet)
            set_outer_border(subset_worksheet, f'B2:B{row_idx - 1}', 'thin')
            set_outer_border(subset_worksheet, f'C2:C{row_idx - 1}', 'thin')
            set_outer_border(subset_worksheet, f'A1:D1', 'thick')
            set_outer_border(subset_worksheet, None, 'thick')

            # Plot frequency bar chart of decision dates
            if 'Decision_Date' in subset:
                # Extract year from each datetime object
                decision_years = [date.year for date in subset['Decision_Date'] if pd.notnull(date)]
                year_counts = Counter(decision_years)
                data_for_chart = list(year_counts.items())
                plot_frequency_bar_chart(workbook, sheet_name, data_for_chart, 'Decision Date Frequency','F1')

if __name__ == "__main__":
    main()
