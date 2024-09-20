import pandas as pd

def compare_excel_files(file1, file2, output_file):
    # Load the Excel files
    excel1 = pd.ExcelFile(file1)
    excel2 = pd.ExcelFile(file2)

    # Create a writer object to write the results to a new Excel file
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    # Iterate through the sheet names of the first file
    for sheet_name in excel1.sheet_names:
        # Check if the sheet exists in both files
        if sheet_name in excel2.sheet_names:
            # Read the data from both files for the given sheet
            df1 = pd.read_excel(file1, sheet_name=sheet_name)
            df2 = pd.read_excel(file2, sheet_name=sheet_name)

            # Find the overlapping data (inner join on all columns)
            overlap = pd.merge(df1, df2, how='inner')

            # If there is overlapping data, write it to the new Excel file
            if not overlap.empty:
                overlap.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                print(f"No overlap found in sheet '{sheet_name}'")

    # Save the output Excel file
    writer.save()
    print(f"Comparison complete. Overlapping data saved in '{output_file}'.")

# Example usage:
file1 = 'file1.xlsx'  # Path to the first Excel file - UPDATE TO REFLECT NAME OF FILE
file2 = 'file2.xlsx'  # Path to the second Excel file - UPDATE TO REFLECT NAME OF FILE
output_file = 'overlap_data.xlsx'  # Path to the output Excel file

compare_excel_files(file1, file2, output_file)