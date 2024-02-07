import pandas as pd

def remove_duplicates(input_file, column_name, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Keep only the first occurrence of each duplicated value in the specified column
    df_no_duplicates = df.drop_duplicates(subset=[column_name], keep='first')

    # Save the result to a new Excel file
    df_no_duplicates.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")


# Example usage:
input_excel_file = 'u:\Temp\Programming\HLBDM\Dataoutput\AWSStage-S2.xlsx'  # Replace with your input file path
output_excel_file = 'u:\Temp\Programming\HLBDM\Dataoutput\Imp-AWSStage-Elem.xlsx'  # Replace with your desired output file path
column_to_check = 'AWS'  # Replace with the actual column name

remove_duplicates(input_excel_file, column_to_check, output_excel_file)
