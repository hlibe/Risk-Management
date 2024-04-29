import pandas as pd
import os
import glob
import re  # For regular expression operations

def extract_rows_and_save_excel(file_keyword, rows_to_extract, output_file_name):
    """
    Load an Excel file containing the specified keyword, extract specified rows,
    rename columns, add a new 'PORT' row, find and append bonds based on various criteria, 
    apply formatting, and save them in a new Excel file.
    """
    # Use glob to find files that contain the keyword in the current directory
    excel_files = glob.glob(os.path.join(os.getcwd(), f'*{file_keyword}*.xlsx'))

    # Check if there is at least one file that matches the criteria
    if not excel_files:
        raise FileNotFoundError(f"No Excel files containing '{file_keyword}' found in the current directory.")
    elif len(excel_files) > 1:
        raise RuntimeError(f"Multiple Excel files containing '{file_keyword}' found. Please specify which one to use.")

    report_df = pd.read_excel('Report_1.xlsx')
    report_df = report_df.drop(report_df.index[0:5])
    # Reset the index after dropping rows
    report_df.reset_index(drop=True, inplace=True)

    # Set the first row as column names
    report_df.columns = report_df.iloc[0]

    # Drop the first row now that it's been set as column names
    report_df = report_df.drop(report_df.index[0])

    # Rename the first column to 'Name'
    report_df.rename(columns={report_df.columns[0]: 'Name'}, inplace=True)

    # Reset the index again
    report_df.reset_index(drop=True, inplace=True)


    # Extract the 'Yield to Worst' column
    yield_to_worst = report_df[['Yield to Worst','Name']].copy()

    # Select the first matching file
    excel_file_path = excel_files[0]

    # Load the Excel file
    df_excel = pd.read_excel(excel_file_path)

    # Rename the first column to 'PortNm'
    df_excel.rename(columns={df_excel.columns[0]: 'PortNm'}, inplace=True)

    # Rename the second column directly to 'SecNum'
    df_excel.rename(columns={df_excel.columns[1]: 'SecNum'}, inplace=True)

    # Extract rows 'Total' and 'Corporate Bond'
    extracted_rows = df_excel[df_excel.iloc[:, 0].isin(rows_to_extract)].copy()

    # Filter bonds based on the 'Name' criteria
    bond_criteria = df_excel['Name'].apply(lambda x: (not x.endswith('RP')) and ('cash' not in x) and ('MULTI' not in x))
    filtered_bonds = df_excel[bond_criteria]

    # Calculate the sumproduct of 'IncCst' and 'CumChgPct', then divide by the sum of 'IncCst'
    cumchg_weighted_avg = (filtered_bonds['IncCst'] * filtered_bonds['CumChgPct']).sum() / filtered_bonds['IncCst'].sum()

    # Calculate the sumproduct of 'IncCst' and 'YTDChgPct', then divide by the sum of 'IncCst'
    ytdchg_weighted_avg = (filtered_bonds['IncCst'] * filtered_bonds['YTDChgPct']).sum() / filtered_bonds['IncCst'].sum()


    # Output filtered_bonds as an Excel file
    filtered_bonds_output_file_name = 'filtered_bonds.xlsx'  # Define the output file name
    filtered_bonds_output_path = os.path.join(os.getcwd(), filtered_bonds_output_file_name)
    filtered_bonds.to_excel(filtered_bonds_output_path, index=False)  # Save the DataFrame to an Excel file

    print(f"Filtered bonds data saved to {filtered_bonds_output_path}")


    # Find the five bonds with the lowest "FiveDayChgPct" and "CumChgPct"
    lowest_five_fiveday = filtered_bonds.nsmallest(5, 'FiveDayChgPct')
    lowest_five_cum = filtered_bonds.nsmallest(5, 'CumChgPct')

    # Append the lowest five bonds based on 'FiveDayChgPct' and 'CumChgPct' to the extracted rows
    extracted_rows = pd.concat([extracted_rows, lowest_five_fiveday, lowest_five_cum])

    # Create the 'PORT' row
    PORT_row = extracted_rows[extracted_rows['PortNm'] == 'Corporate Bond'].copy()
    PORT_row['PortNm'] = 'PORT'  # Set the name of the row to 'PORT'
    PORT_row['Lev'] = PORT_row['Lev'] - 100  # Decrease 'Lev' by 100 for 'PORT' row
    
    # Place the 'PORT' row before the rows 'Total'
    extracted_rows = pd.concat([PORT_row, extracted_rows[extracted_rows['PortNm'] != 'Corporate Bond']])

    # Now you can safely use 'PortNm' column
    extracted_rows.loc[extracted_rows['PortNm'] == 'PORT', 'Name'] = 'PROP2'

    # PORTine extracted_rows and yield_to_worst using the 'Name' column
    extracted_rows = pd.merge(extracted_rows, yield_to_worst, on='Name', how='left')

    # Find the position of the 'YldAtCst' column
    yldatcst_position = extracted_rows.columns.get_loc('YldAtCst')

    # Move the 'Yield to Worst' column to be just after 'YldAtCst'
    yield_to_worst_column = extracted_rows.pop('Yield to Worst')
    extracted_rows.insert(yldatcst_position + 1, 'Yield to Worst', yield_to_worst_column)

    # Remove the rows 'Total' and 'Corporate Bond'
    extracted_rows = extracted_rows[~extracted_rows['PortNm'].isin(['Total', 'Corporate Bond'])]

    # Remove the 'Data Export Restricted' column if it exists
    if 'Data Export Restricted' in extracted_rows.columns:
        extracted_rows.drop('Data Export Restricted', axis=1, inplace=True)

    # Count Investment Grade bonds
    investment_grade_bonds = len(df_excel[df_excel['BBRt'].isin(['AAA', 'AA+', 'AA', 'AA-', 'A+', 'A', 'A-', 'BBB+', 'BBB', 'BBB-'])])

    # Count Non-IG bonds
    non_ig_bonds = len(df_excel[df_excel['BBRt'].isin(['BB+', 'BB', 'BB-', 'B+', 'B', 'B-', 'CCC+', 'CCC', 'CCC-', 'CC', 'C', 'DDD', 'DD', 'D'])])

    # Count Non-Rated bonds
    non_rated_bonds = len(df_excel[df_excel['BBRt'].isin(['N.A.', 'NR'])])

    # Add the new columns with initial values
    extracted_rows['Investment Grade'] = 0
    extracted_rows['Non-IG'] = 0
    extracted_rows['Non-Rated'] = 0



    # Update values for the row with 'Name' = 'PROP2'
    prop2_index = extracted_rows[extracted_rows['Name'] == 'PROP2'].index
    if not prop2_index.empty:
        extracted_rows.at[prop2_index[0], 'Investment Grade'] = investment_grade_bonds
        extracted_rows.at[prop2_index[0], 'Non-IG'] = non_ig_bonds
        extracted_rows.at[prop2_index[0], 'Non-Rated'] = non_rated_bonds

    # Now you can safely use 'PortNm' column
    extracted_rows.loc[extracted_rows['PortNm'] == 'PORT', 'Name'] = 'PROP2'

    # Update the 'CumChgPct' and 'YTDChgPct' for 'PROP2'
    prop2_index = extracted_rows[extracted_rows['Name'] == 'PROP2'].index
    if not prop2_index.empty:
        extracted_rows.at[prop2_index[0], 'CumChgPct'] = cumchg_weighted_avg
        extracted_rows.at[prop2_index[0], 'YTDChgPct'] = ytdchg_weighted_avg
    
    extracted_rows.loc[extracted_rows['Name'] != 'PROP2', 'Lev'] = pd.NA
    #extracted_rows.loc[extracted_rows['Name'] != 'Yield to Worst', 'Lev'] = pd.NA

    # Fill empty cells with 'N.A.' in the row where 'Name' is 'PROP2'
    prop2_row_index = extracted_rows[extracted_rows['Name'] == 'PROP2'].index
    if not prop2_row_index.empty:
        extracted_rows.loc[prop2_row_index, :] = extracted_rows.loc[prop2_row_index, :].fillna('N.A.')

    # Replace zeros with 'N.A.' in specific columns except for the row where 'Name' is 'PROP2'
    cols_to_update = ['Investment Grade', 'Non-IG', 'Non-Rated']
    prop2_row_index = extracted_rows[extracted_rows['Name'] == 'PROP2'].index

    # Check if PROP2 row exists to avoid modifying it
    if not prop2_row_index.empty:
        prop2_index = prop2_row_index[0]
        for col in cols_to_update:
            extracted_rows.loc[(extracted_rows.index != prop2_index) & (extracted_rows[col] == 0), col] = 'N.A.'
    else:
        # If there's no PROP2 row, modify all zeros in the specified columns
        for col in cols_to_update:
            extracted_rows.loc[extracted_rows[col] == 0, col] = 'N.A.'
    print("Column names after setting the first row as column names:", extracted_rows.columns.tolist())
    # Rename the columns to the specified Chinese names
    column_name_mappings = {
        'PortNm': '组合',
        'SecNum': '券数',
        'Name': '券名',
        'ChNm': '中文名',
        'PrtCoNm': '母公司名',
        'UltPrtCoNm': '最终母公司',
        'TotMktVal': '总市值',
        'YldAtCst': '成本收益率(%)',
        'Yield to Worst': '最差收益率(%)',
        'BBRt': '信用评级',
        'IncCst': '初始成本',
        'Lev': '杠杆率(%)',
        'Pos': '持仓',
        'FiveDayChgPct': '五日变动百分比',
        'CumChgPct': '累积变动百分比',
        'YTDChgPct': '年初至今盈亏(%)',
        'Tot_PLoS': '证券总盈亏',
        'YTD_PLoSC': '年初至今总盈亏',
        'YTD_PLoS': '年初至今已实现盈亏',
        'YTD_UPLoS': '年初至今未实现盈亏',
        'YTD_Carry': '年初至今累计利息',
        'Investment Grade': '投资级',
        'Non-IG': '非投资级',
        'Non-Rated': '无评级'
    }

    extracted_rows.rename(columns=column_name_mappings, inplace=True)

    # Save the extracted rows to an Excel file with xlsxwriter engine
    output_path = os.path.join(os.getcwd(), output_file_name)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    extracted_rows.to_excel(writer, sheet_name='Sheet1', index=False)

    # Get the xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Define the red format
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

    # Apply red format to cells where '五日变动百分比' < -5 and '累积变动百分比' < -8
    fiveday_column = extracted_rows.columns.get_loc('五日变动百分比')  # Adjusting for Excel's 1-based indexing
    cumchg_column = extracted_rows.columns.get_loc('累积变动百分比')
    for idx, row in extracted_rows.iterrows():
        if row['五日变动百分比'] < -5:
            worksheet.write(idx , fiveday_column, row['五日变动百分比'], red_format)
        if row['累积变动百分比'] < -8:
            worksheet.write(idx , cumchg_column, row['累积变动百分比'], red_format)

    # Close the Pandas Excel writer and output the Excel file
    writer.close()

    print(f"Extracted data saved to {output_path}")

# Specify the rows to extract and the name of the output file
rows_to_extract = ['Total', 'Corporate Bond']
output_file_name = 'port_monitor.xlsx'

# Call the function
extract_rows_and_save_excel('grid', rows_to_extract, output_file_name)
