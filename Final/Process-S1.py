import pandas as pd

def process_excel(input_file, output_file_kp, output_file_tp, column_name):
    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Filter rows containing "KP" in the specified column
    kp_df = df[df[column_name].str.strip().str.startswith("KP", na=False)]
    
    # Replace blanks with "_" and remove leading/trailing spaces
    #kp_df[column_name] = kp_df[column_name].str.strip().replace("", "_")
    kp_df[column_name] = kp_df[column_name].str.strip().replace("", "_").replace(" ", "_")
    
    # Save the "KP" results to a new Excel file
    kp_df.to_excel(output_file_kp, index=False)

    # Filter rows containing "TP" in the specified column
    tp_df = df[df[column_name].str.strip().str.contains("TP", na=False)]

    # Replace blanks with "_" and remove leading/trailing spaces
    tp_df[column_name] = tp_df[column_name].str.strip().replace("", "_")

    # Save the "TP" results to a new Excel file
    tp_df.to_excel(output_file_tp, index=False)

if __name__ == "__main__":
    # Specify your input and output file paths, and the column name
    input_excel = "u:\Temp\Programming\HLBDM\Data\Processes-ohne-Beschr.xlsx"
    output_excel_kp = "u:\Temp\Programming\HLBDM\Dataoutput\output_file_kp.xlsx"
    output_excel_tp = "u:\Temp\Programming\HLBDM\Dataoutput\output_file_tp.xlsx"
    target_column = "AT_DESCRIPTION"

    # Process the Excel file
    process_excel(input_excel, output_excel_kp, output_excel_tp, target_column)

    print("Processing complete. Check the output files:", output_excel_kp, output_excel_tp)
