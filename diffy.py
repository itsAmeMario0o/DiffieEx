import pandas as pd

def compare_excel_files(file1, file2, output_file):
    # Load the Excel files
    excel1 = pd.ExcelFile(file1)
    excel2 = pd.ExcelFile(file2)

    # Create a writer object to write the results to a new Excel file
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    # Iterate through all sheet names in file1
    for sheet_name1 in excel1.sheet_names:
        # Read the data from the current sheet in file1
        df1 = pd.read_excel(file1, sheet_name=sheet_name1)
        
        # Now iterate through all sheet names in file2
        for sheet_name2 in excel2.sheet_names:
            # Read the data from the current sheet in file2
            df2 = pd.read_excel(file2, sheet_name=sheet_name2)

            # Find common columns
            common_columns = df1.columns.intersection(df2.columns)

            if not common_columns.empty:
                # Find the overlapping data (inner join on common columns)
                overlap = pd.merge(df1, df2, how='inner', on=common_columns.tolist())

                # If there is overlapping data, write it to the new Excel file
                if not overlap.empty:
                    output_sheet_name = f"{sheet_name1}_vs_{sheet_name2}"
                    output_sheet_name = output_sheet_name[:31]  # Truncate if necessary
                    
                    overlap.to_excel(writer, sheet_name=output_sheet_name, index=False)
                else:
                    print(f"No overlap found for '{sheet_name1}' vs '{sheet_name2}'.")
            else:
                print(f"No common columns to merge on between '{sheet_name1}' and '{sheet_name2}'.")

    # Save and close the output Excel file
    writer.close()
    print(f"Comparison complete. Overlapping data saved in '{output_file}'.")

# Example usage:
file1 = 'DOC1.xlsx'  # Path to the first Excel file
file2 = 'DOC2.xlsx'  # Path to the second Excel file
output_file = 'overlap_data.xlsx'  # Path to the output Excel file

compare_excel_files(file1, file2, output_file)