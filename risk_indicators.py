import pandas as pd
import os
import glob

def load_csv_and_extract_data(file_keyword, column_mappings, rows):
    """
    Load CSV file containing the specified keyword and extract data for the specified columns and rows.
    """
    # Use glob to find files that contain the keyword in the current directory
    csv_files = glob.glob(os.path.join(current_directory, f'*{file_keyword}.csv'))

    # Check if there is at least one file that matches the criteria
    if not csv_files:
        raise FileNotFoundError(f"No CSV files containing '{file_keyword}' found in the current directory.")
    elif len(csv_files) > 1:
        raise RuntimeError(f"Multiple CSV files containing '{file_keyword}' found. Please specify which one to use.")

    # Select the first matching file
    csv_file_path = csv_files[0]

    # Load the data section of the CSV file
    with open(csv_file_path, 'r') as file:
        content = file.readlines()
    data_start_index = [i for i, line in enumerate(content) if 'DATA_START' in line][0] + 1
    df_data = pd.read_csv(csv_file_path, skiprows=data_start_index, delimiter='|')

    # Extracting the specified values from the specified columns
    for row in rows:
        if row in df_data['CNCB_TEST:node_name'].values:
            for new_col, csv_col in column_mappings.items():
                if csv_col in df_data.columns:
                    value = df_data[df_data['CNCB_TEST:node_name'] == row][csv_col].values[0]
                    new_df.at[row, new_col] = value
                else:
                    print(f"Column '{csv_col}' not found in the CSV file.")

# Get the current directory of the script
current_directory = os.getcwd()

# Define the rows and columns for the new DataFrame
rows = ["PROP2", "SP1", "SP4", "SP6"]
columns = ["IR01", "CR01", "VaR", "Duration"]

# Create the new DataFrame with NaN values initially
new_df = pd.DataFrame(index=rows, columns=columns)

# Process the file containing 'Duration_node_CNCB_TEST'
duration_column_mappings = {"Duration": "STATISTICS:Duration Parallel Spot( SHIFT_LEVEL:1bp TENOR:Total )"}
load_csv_and_extract_data('Duration_node_CNCB_TEST', duration_column_mappings, rows)

# Process the file containing 'Greeks_node_CNCB_TEST'
greek_column_mappings = {
    "IR01": "STATISTICS:IR01 Parallel (zero)( SHIFT_LEVEL:1bp TENOR:Total )",
    "CR01": "STATISTICS:CR01 Parallel( SHIFT_LEVEL:1bp TENOR:Total )"
}
load_csv_and_extract_data('Greeks_node_CNCB_TEST', greek_column_mappings, rows)

# Process the file containing 'HVAR_node_CNCB_TEST'
var_column_mappings = {"VaR": "STATISTICS:1 Year 1D99 VaR"}
load_csv_and_extract_data('HVAR_node_CNCB_TEST', var_column_mappings, rows)

# Define the path for the new Excel file (in the same directory)
output_excel_path = os.path.join(current_directory, 'risk_indicators.xlsx')

# Saving the new DataFrame as an Excel file
new_df.to_excel(output_excel_path)