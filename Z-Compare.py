import pandas as pd

# Function to compare columns and update sheet three
def compare_and_update(sheet1, column1_sheet1, sheet2, column2_sheet2, sheet3):
    # Read Excel sheets
    df_sheet1 = pd.read_excel(sheet1)
    df_sheet2 = pd.read_excel(sheet2)

    # Select specific columns
    column1_data_sheet1 = df_sheet1[column1_sheet1]
    column2_data_sheet2 = df_sheet2[column2_sheet2]

    # Find values present in sheet1 and not in sheet2
    source_values = column1_data_sheet1[~column1_data_sheet1.isin(column2_data_sheet2)]

    # Find values present in sheet2 and not in sheet1
    target_values = column2_data_sheet2[~column2_data_sheet2.isin(column1_data_sheet1)]

    # Create a new DataFrame for sheet3
    df_sheet3 = pd.DataFrame(columns=["Value", "Type"])

    # Update sheet3 with source values
    df_sheet3.loc[df_sheet3.shape[0]:, "Value"] = source_values
    df_sheet3.loc[df_sheet3.shape[0] - source_values.shape[0]:, "Type"] = "source"

    # Concatenate sheet3 with target values
    df_sheet3 = pd.concat([df_sheet3, pd.DataFrame({"Value": target_values, "Type": "target"})])

    # Save the result to sheet3
    df_sheet3.to_excel(sheet3, index=False)

# Example usage
input_file1_path = "u:\Temp\Programming\HLBDM\Data\ACH-AWS-S1.xlsx"
input_file2_path = "u:\Temp\Programming\HLBDM\Dataoutput\SchnittSt-S1.xlsx"
output_file_path = "u:\Temp\Programming\HLBDM\Dataoutput\Diff.xlsx"
column_index1 = 'Unnamed: 11'  # Index of the  column in the first sheet
column_index2 = 'Allgemein'  # Index of the  column in the second sheet    

compare_and_update(input_file1_path, column_index1, input_file2_path, column_index2, output_file_path)
