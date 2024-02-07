# This script takes in all the DB Versions from the CMDB, cuts away all minor versions,
# and makes this a key for an Index for Archi.
# In: ACH-Software-DB which comes from CMDB-DBmanagementsysteme
# The output is in step one converted manually to an elements file for the DB-SW
# Then you run it on the Instances to get the releation DBSW to Instance

import pandas as pd
import re

def drop_parentheses(cell_value):
    # Use regular expression to remove parentheses and content inside them
    return re.sub(r'\([^)]*\)', '', cell_value)

def replace_core_based_licensing(cell_value):
    # Replace "Core-based" with "Corebased"
    return cell_value.replace("Core-based", "Corebased")

def drop_after_third_character_after_first_hyphen(cell_value):
    # Find the index of the first hyphen
    first_hyphen_index = cell_value.find('-')

    # Check if hyphen is present and there are at least three characters after it
    if first_hyphen_index != -1 and len(cell_value) > first_hyphen_index + 3:
        # Find the index of the third character after the first hyphen
        third_character_index = first_hyphen_index + 3

        # Return the substring up to the third character after the first hyphen
        return cell_value[:third_character_index]

    return cell_value

def drop_after_value_and_next_n_chars(cell_value, value_to_drop, n):
    # Check if the specified value is present and retain characters up to n characters after its occurrence
    value_index = cell_value.find(value_to_drop)
    if value_index != -1 and len(cell_value) > value_index + len(value_to_drop) + n:
        return cell_value[:value_index + len(value_to_drop) + n]
    return cell_value

def process_columns(df, columns_to_process):
    for col in columns_to_process:
        # Apply the drop_parentheses function to the specified column
        df.iloc[:, col] = df.iloc[:, col].apply(drop_parentheses)

        # Replace "Core-based Licensing " with "Corebased Licensing" in the specified column
        df.iloc[:, col] = df.iloc[:, col].apply(replace_core_based_licensing)

        # Apply the drop_after_third_character_after_first_hyphen function to the specified column
        df.iloc[:, col] = df.iloc[:, col].apply(drop_after_third_character_after_first_hyphen)

        # Replace with the value and the next 3 characters in the specified column
        df.iloc[:, col] = df.iloc[:, col].apply(lambda x: drop_after_value_and_next_n_chars(x, 'Oracle', 3))

        # Replace banks with "_" and colons with "-"
        df.iloc[:, col] = df.iloc[:, col].apply(replace_banks_and_colons)

    return df

def replace_banks_and_colons(cell_value):
    # Replace banks with "_"
    cell_value = cell_value.replace(' ', '_')

    # Replace colons with "-"
    cell_value = cell_value.replace(':', '-')

    # Replace colons with "-"
    cell_value = cell_value.replace('\\', '-')

    return cell_value


def main(input_excel_file_path, output_excel_file_path, columns_to_process):
    # Read the Excel file
    df = pd.read_excel(input_excel_file_path, sheet_name='Sheet1')

    # Process specified columns
    df = process_columns(df, columns_to_process)

    # Drop duplicate cells
#    df = df.drop_duplicates()

    # Save the modified DataFrame to a new Excel file
    df.to_excel(output_excel_file_path, index=False)

if __name__ == "__main__":
    # Define file paths
    input_excel_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\DB-S1.xlsx'  
    output_excel_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\RelDBIns2SWDB.xlsx'  

    # Define column indices to process
    columns_to_process = [3,]  # Replace with the actual column indices you want to process

    # Call the main function
    main(input_excel_file_path, output_excel_file_path, columns_to_process)

