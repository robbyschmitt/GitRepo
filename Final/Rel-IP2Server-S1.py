import pandas as pd

def process_and_compare(input_file1, input_file2, output_file, column_to_process1, column_to_process2, column_to_compare1):
    # Compare the Server List from Read the first Excel sheet with the servers from Nexpose and get the IPs
    df1 = pd.read_excel(input_file1)

    # Convert values in the specified column to strings and then process
    df1[column_to_process1] = df1[column_to_process1].astype(str)
    df1[column_to_process1] = df1[column_to_process1].apply(lambda x: x.split('.')[0])
    df1[column_to_process1] = df1[column_to_process1].str.upper()

    # Read the second Excel sheet
    df2 = pd.read_excel(input_file2)

    # Remove "-Node" from values in the second column of the second sheet
    df2[column_to_process2] = df2[column_to_process2].astype(str)
    df2[column_to_process2] = df2[column_to_process2].str.replace("-Node", "")

    # Filter rows in the first sheet based on matching values in the second sheet
    matching_rows = df1[df1[column_to_compare1].isin(df2[column_to_process2])]

    # Save the matching rows to a new Excel sheet
    matching_rows.to_excel(output_file, index=False)

def format(inputfile, outputfile2):
    df2 = pd.read_excel(inputfile)
    # Delete superflous columns
    #columns_to_drop = list([0, 1, 2, 4,]) + list[range(5-44)] 
    df2 = df2.drop(df2.columns[[0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11] + list(range(13, 45))], axis=1)


    df2.to_excel(outputfile2, index=False)


if __name__ == "__main__":
    input_file1_path = r"u:\Temp\Programming\HLBDM\Data\HLB-Assets.xlsx"
    input_file2_path = r"u:\Temp\Programming\HLBDM\Dataoutput\Server-S1-Elem.xlsx"
    output_file_path = r"u:\Temp\Programming\HLBDM\Dataoutput\SRV-knownIP.xlsx"
    outputfile2 = r"u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2Sub.xlsx"
    column_to_process1 = "Asset Name"  # Replace with the actual column name from the first sheet
    column_to_process2 = "Server"  # Replace with the actual column name from the second sheet
    column_to_compare1 = "Asset Name"  # Replace with the actual column name from the first sheet

    process_and_compare(input_file1_path, input_file2_path, output_file_path, column_to_process1, column_to_process2, column_to_compare1)
    format(output_file_path, outputfile2)