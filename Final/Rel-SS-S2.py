import pandas as pd


def process_A2B(input_file, output_fileA2B, search_column_name, Ssid, Source, Target):
    # Take the modified SS File, search for all A->B SS and replace the logic with two rows A->SS and SS->B
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # First replace the IS<xxx> with IS<xxx>_Produktion to bind the SS to the production Stage
    df.iloc[:, 2] = df.iloc[:, 2] + '_Produktion'
    df.iloc[:, 3] = df.iloc[:, 3] + '_Produktion'

    # Find rows where the specified string is present in the specified column
    mask = df[search_column_name].str.contains("A->B")
    
    df = df[mask].copy() #Erase all the other cells

    # Create a new DataFrame with rows that match the condition
    new_rows = df[mask].copy()
    new_rows_copy = new_rows.copy()

    # Swap the contents of two specific columns in the new rows
    new_rows[Target]  = new_rows[Ssid].copy()
    
    # Create a second copy of the new rows with a different swap
    new_rows_copy[Source], new_rows_copy[Source] = new_rows_copy[Target].copy(), new_rows_copy[Ssid].copy()

    # Drop the original rows from the DataFrame
    df.drop(df[mask].index, inplace=True)

    # Concatenate the new DataFrame and the copies
    result_df = pd.concat([df, new_rows, new_rows_copy], ignore_index=True)

    # Write the result to a new Excel file
    result_df.to_excel(output_fileA2B, index=False)


def process_B2A(input_file, output_fileB2A, search_column_name, Ssid, Source, Target):
    # Take the modified SS File, search for all B>A SS and replace the logic with two rows B->SS and SS->A
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # First replace the IS<xxx> with IS<xxx>_Produktion to bind the SS to the production Stage
    df.iloc[:, 2] = df.iloc[:, 2] + '_Produktion'
    df.iloc[:, 3] = df.iloc[:, 3] + '_Produktion'


    # Find rows where the specified string is present in the specified column
    mask = df[search_column_name].str.contains("A<-B")

    df = df[mask].copy() #Erase all the other cells

    # Create a new DataFrame with rows that match the condition
    B2ss = df[mask].copy()
    Ss2a = df[mask].copy()

    # Swap the contents of two specific columns in the new rows
    B2ss[Source], B2ss[Target], = B2ss[Source].copy(), B2ss[Ssid].copy()
      
    # Create a second copy of the new rows with a different swap
    Ss2a[Source], Ss2a[Target]  = Ss2a[Ssid].copy(), Ss2a[Target].copy()

    # Drop the original rows from the DataFrame
    df.drop(df[mask].index, inplace=True)

    # Concatenate the new DataFrame and the copies
    result_df = pd.concat([df, B2ss, Ss2a], ignore_index=True)

    # Write the result to a new Excel file
    result_df.to_excel(output_fileB2A, index=False)

# Take the modified SS File, search for all B>A SS and replace the logic with two rows B->SS and SS->A
def process_A2B2A(input_file, output_fileA2B2A, search_column_name, Ssid, Source, Target):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # First replace the IS<xxx> with IS<xxx>_Produktion to bind the SS to the production Stage
    df.iloc[:, 2] = df.iloc[:, 2] + '_Produktion'
    df.iloc[:, 3] = df.iloc[:, 3] + '_Produktion'

    # Find rows where the specified string is present in the specified column
    mask = df[search_column_name].str.contains("A<->B")

    df = df[mask].copy() #Erase all the other cells

    # Create a new DataFrame with rows that match the condition
    new_rows1 = df[mask].copy()
    new_rows2 = new_rows1.copy()
    new_rows3 = df[mask].copy()
    new_rows4 = df[mask].copy()

    # Swap the contents of two specific columns in the new rows
    new_rows1[Source]  = new_rows1[Ssid].copy()
    
    # Create a second copy of the new rows with a different swap
    new_rows2[Target] = new_rows2[Ssid].copy()

    # Create a second copy of the new rows with a different swap
    new_rows3[Target], new_rows3[Source] = new_rows3[Ssid].copy(), new_rows3[Target].copy()

    # Create a second copy of the new rows with a different swap
    new_rows4[Source], new_rows4[Target] = new_rows4[Ssid].copy(), new_rows4[Source].copy()

    # Drop the original rows from the DataFrame
    df.drop(df[mask].index, inplace=True)

    # Concatenate the new DataFrame and the copies
    result_df = pd.concat([df, new_rows1, new_rows2, new_rows3, new_rows4], ignore_index=True)

    # Write the result to a new Excel file
    result_df.to_excel(output_fileA2B2A, index=False)

def concat(output_fileA2B, output_fileB2A, output_fileA2B2A):
    A2b = pd.read_excel(output_fileA2B)
    B2a = pd.read_excel(output_fileB2A)
    A2b2a = pd.read_excel(output_fileA2B2A)

# Concatenate the new DataFrame and the copies
    result_df = pd.concat([A2b, B2a, A2b2a, ], ignore_index=False)
    
    # Write the result to a new Excel file

    result_df.to_excel(output_filefinal, index=False)

if __name__ == "__main__":
    # Replace these values with your actual file paths and column names
    input_file = "u:\Temp\Programming\HLBDM\Dataoutput\SchnittSt-S1.xlsx"
    output_fileA2B = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SS1-Rel.xlsx"
    output_fileB2A = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SS2-Rel.xlsx"
    output_fileA2B2A = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SS3-Rel.xlsx"
    output_filefinal = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SSFin-Rel.xlsx"
    search_column_name = "Unnamed: 13"
    Ssid = "Unnamed: 7"
    Source = "Unnamed: 9"
    Target = "Unnamed: 11"

    process_A2B(input_file, output_fileA2B, search_column_name, Ssid, Source, Target)
    process_B2A(input_file, output_fileB2A, search_column_name, Ssid, Source, Target)
    process_A2B2A(input_file, output_fileA2B2A, search_column_name, Ssid, Source, Target)
    concat(output_fileA2B, output_fileB2A, output_fileA2B2A)
    # Write the result of all steps to a new Excel file
    
    print("Script completed successfully.")
