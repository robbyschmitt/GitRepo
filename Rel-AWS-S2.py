import pandas as pd

def process_excel_file(input_file, output_filtered_file, value_to_find):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Filter rows based on the condition in column 6
    filtered_df = df[df.iloc[:, 5] == value_to_find]

    # Update values in column 1 with '-bid' added as a suffix
    filtered_df.iloc[:, 0] = filtered_df.iloc[:, 0].astype(str) + '-bid'

    mask = df.iloc[:, 5] == value_to_find
    filtered_df.loc[mask, [df.columns[2], df.columns[3]]] = df.loc[mask, [df.columns[3], df.columns[2]]].values


    # Write the filtered DataFrame to a new Excel sheet
    filtered_df.to_excel(output_filtered_file, index=False)


if __name__ == "__main__":
    # Replace these values with your actual file paths and desired value
    input_excel_file = 'u:\Temp\Programming\HLBDM\Dataoutput\AWS-Imp-Step3.xlsx'
    output_filtered_file = 'u:\Temp\Programming\HLBDM\Dataoutput\AWS-Imp-Step3a.xlsx'
    value_to_find = "A<->B"

    process_excel_file(input_excel_file, output_filtered_file, value_to_find)
    
    print(f"Excel file processed and saved as {output_filtered_file}")