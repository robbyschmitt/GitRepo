import pandas as pd
import openpyxl

def process_excel_file(input_file,input_file2, output_file, output_file2, columnTP, columnKP):
    # Read the Excel file TP
    df1 = pd.read_excel(input_file, engine='openpyxl')
    # Read the excel KP to compare
    df2 = pd.read_excel(input_file2, engine='openpyxl')

    # Extract characters between the first and second underscores into a new column
    df1['TP'] = df1[columnTP].str.split('_', n=2).str[1]

    # Add the KP_ to match the KP ID
    df1['TP'] = 'KP_' + df1['KP']

    merged_df = pd.merge(df1, df2, how='inner', left_on=columnKP, right_on=columnTP)

    # Extract the matching values
   # matching_values = merged_df[columnKP]

    # Create a new DataFrame with the matching values
    #result_df = pd.DataFrame({columnKP: matching_values})

    # Save the modified DataFrame to a new Excel file
    merged_df.to_excel(output_file, index=False, engine='openpyxl')
    df1.to_excel(output_file2, index=False, engine='openpyxl')

def KP2TP(filename):
    # Load the workbook
    wb = openpyxl.load_workbook(filename)
    # Select the active sheet
    sheet = wb.active

    # Iterate over rows starting from the second row
    for row in sheet.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            # Replace "TP_" with "KP_"
            cell.value = cell.value.replace("TP_", "KP_")
            # Find the index of the second "_" character
            index = cell.value.find("_", cell.value.find("_") + 1)
            # If the second "_" exists, keep only characters before it
            if index != -1:
                cell.value = cell.value[:index]

    # Save the workbook
    wb.save(filename)

if __name__ == "__main__":
    # Specify the input and output file paths and the column to process
    filename = 'u:\Temp\Programming\HLBDM\Data\ACH-TeilProzesse.xlsx'
    input_file_path = 'u:\Temp\Programming\HLBDM\Data\ACH-TeilProzesse.xlsx' 
    input_file_path2 = 'u:\Temp\Programming\HLBDM\Data\ACH-KernProzesse.xlsx'
    out_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\zzz.xlsx' 
    out_file_path2 = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-TP2KP.xlsx'
    columnTP = 'KP' 
    columnKP =  'AT_DESCRIPTION'

    # Process the Excel file
    KP2TP(filename)

    print(f"Excel file processed and saved to {out_file_path2}")
