import pandas as pd


def add_suffix_to_columns(input_file, output_file, column_to_read, column_to_update_1, column_to_update_2):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Add a new column with the values from the specified column as a suffix
    df[column_to_update_1] = df[column_to_update_1].astype(str) + '_' + df[column_to_read].astype(str) 
    df[column_to_update_2] = df[column_to_update_2].astype(str) + '_' + df['Server'].astype(str)

    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

def remove_duplicates(output_file_path, column_to_update_1_name, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(output_file_path)

    # Keep only the first occurrence of each duplicated value in the specified column
    df_no_duplicates = df.drop_duplicates(subset=[column_to_update_1_name], keep='first')

    # Save the result to a new Excel file
    df_no_duplicates.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")


# Specify the input and output file paths, and column names
input_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\AWSStage-Step2.xlsx'
output_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\AWSStage-Step3.xlsx'
output_file = 'u:\Temp\Programming\HLBDM\Dataoutput\AWSStage-Step4.xlsx'
column_to_read_name = 'Umgebung'
column_to_update_1_name = 'Nr'
column_to_update_2_name = 'Umgebung'

# Call the function with the specified parameters
add_suffix_to_columns(input_file_path, output_file_path, column_to_read_name, column_to_update_1_name, column_to_update_2_name)
remove_duplicates(output_file_path, column_to_update_1_name, output_file)