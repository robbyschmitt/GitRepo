import pandas as pd

def process_and_save(input_file, output_file, column_to_process):
    # Read the Excel sheet
    df = pd.read_excel(input_file)

    # Process values in the specified column
    df[column_to_process] = df[column_to_process].astype(str)
    df[column_to_process] = df[column_to_process].str.replace(".hlb.helaba.de", "")

    # Save the resulting dataframe as a new sheet in the same Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ProcessedSheet', index=False)

if __name__ == "__main__":
    input_file_path = r"u:\Temp\Programming\HLBDM\Data\ACH-Assets-Nexpose.xlsx"
    output_file_path = r"u:\Temp\Programming\HLBDM\Data\ACH-Assets-Nexpose-cleansrv.xlsx"
    column_to_process = "Asset Name"  # Replace with the actual column name

    process_and_save(input_file_path, output_file_path, column_to_process)
