import pandas as pd
from utils import get_executable_path
from LKK_Excel_Generator.data_processing import extract_decisions_and_cases
from LKK_Excel_Generator.excel_formatting import apply_excel_formatting

def main():
    # Print the current working directory using a utility function
    print("Current Working Directory:", get_executable_path())

    # Input: Request user to enter the name of the input and output Excel files
    input_file_path = input("Please enter the name of the excel file (without extension): ")
    output_file_path = input("Please enter the name of the saved excel file (without extension): ")

    # Append '.xlsx' extension to the file names
    input_file_path = input_file_path + '.xlsx'
    output_file_path = output_file_path + '.xlsx'

    # Try to read the Excel file. If an error occurs, print the error message and return
    try:
        df = pd.read_excel(get_executable_path() + '/' + input_file_path)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    # Data Processing: Select specific columns and convert dates to date-only format
    df = df.iloc[:, [0, 2]]  # Adjust to select the relevant columns
    df.iloc[:, 1] = pd.to_datetime(df.iloc[:, 1]).dt.date

    # Extract decisions and cases from the data
    decisions = []
    for _, row in df.iterrows():
        decision_name = row.iloc[0]
        decision_date = row.iloc[1]
        cases = extract_decisions_and_cases(decision_name)
        cases_string = ' & '.join(cases)
        decisions.append([None, decision_name, decision_date, cases_string])  # Prepares data for new DataFrame

    # Create a new DataFrame with formatted data and sort it by date
    all_decisions_df = pd.DataFrame(decisions, columns=[None, 'Filename', 'Decision_Date', 'Cases_Numbers'])
    all_decisions_df = all_decisions_df.sort_values(by=['Decision_Date'])

    # Write the processed data to a new Excel file with formatting
    with pd.ExcelWriter(get_executable_path() + '/' + output_file_path, engine='openpyxl') as writer:
        all_decisions_df.to_excel(writer, index=False, sheet_name='Cases_Numbers')
        apply_excel_formatting(writer, all_decisions_df)  # Apply custom formatting to the Excel file

if __name__ == "__main__":
    main()
