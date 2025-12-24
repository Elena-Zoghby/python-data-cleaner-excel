import pandas as pd
import os


def run_data_cleaner(input_filename='corrupt-sheet.xlsx', output_filename='clean-sheet.csv'):
    """
    Loads a messy CSV file, cleans it, and exports the standardized data.
    """
    print(f"---Starting Data Cleaning Process for {input_filename}---")

    # check if file exists before attempting to load
    if not os.path.exists(input_filename):
        print(f"ERROR: Input file not found.")
        print("Please ensure the input file is in the same folder as this script.")
        return
    try:
        # Load the csv file into a pandas dataframe
        df = pd.read_excel(input_filename)
    except Exception as e:
        print(f"ERROR: Failed to read CSV file. Reason: {e}")
        return
    # record initial state for the final report
    initial_rows = len(df)
    print(f"Initial row count: {initial_rows}")

    # fixing email addresses and duplicates
    print("\n Cleaning and removing duplicates...")
    # lowercase letters in emails for consistency
    df['Email Address'] = df['Email Address'].astype(str).str.lower()
    # remove duplicate emails
    df = df.drop_duplicates(subset=['Email Address'], keep='first')
    rows_after_duplicates = len(df)
    removed_duplicates = initial_rows - rows_after_duplicates
    print(f"Successfully removed {removed_duplicates} duplicate rows")

    # standardizing name and text
    print("\n Standardizing names and text fields...")
    df['Customer Name'] = df['Customer Name'].astype(str).str.title()

    # cleaning up service column
    df['Service'] = df['Service'].astype(str).str.strip().str.title()

    # numerical data
    print("\n Cleaning numerical data amd handling missing values...")
    df['Order Value'] = (df['Order Value'].astype(str).str.replace(
        r'[$]', '', regex=True).str.replace(',', '', regex=False))

    # handling missing values and convert to final number type
    # fill missing cells with 0.0
    df['Order Value'] = df['Order Value'].fillna(0.0)
    try:
        df['Order Value'] = df['Order Value'].astype(float)
    except ValueError as e:
        print(
            f"WARNING: Could not convert Order Value to number. Check for unhandled characters. Error: {e}")

    # export
    print("\n Exporting clean data and final report")
    df.to_csv(output_filename, index=False)

    print("Data Cleaning Complete ")
    print(f"Output file: {output_filename}")
    print("The file is ready!")


if __name__ == "__main__":
    run_data_cleaner()
