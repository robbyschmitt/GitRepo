# Imports 

import pandas as pd

def formatSheet(excel_file2_path, OutputExcel1):
    """Format th original AWSDB-Server Input sheet so we can work with it  """
# Read the Excel file
    input_file = excel_file2_path
    df = pd.read_excel(input_file)

# Unmerge the cells in the first row and the first column
    df.iloc[0, :] = df.iloc[0, :].ffill()
    df.iloc[:, 0] = df.iloc[:, 0].ffill()

# Drop the first row 
    df = df.iloc[0:, ]


# Delete specified columns (first, third, sixth, seventh, and all the rest)
    columns_to_drop = list([0, 2, 6, 7])
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Replace blanks, "[" and "]" with "_"
#    df.iloc[:, 1] = df.iloc[:, 1].astype(str).replace(to_replace={' ': '_', '\[': '',  '\]': ''}, regex=True)

# Write the results into a different Excel file
    df.to_excel(OutputExcel1, index=False, header=False)

def multiplestages(InputExcel, output_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(InputExcel)
    
    # Create an empty DataFrame for the output
    output_df = pd.DataFrame(columns=df.columns)
    
    # Iterate through each row in the DataFrame to search for | 
    for _, row in df.iterrows():

        # Get the value in the second column Name of Server with [XXX]
        cell_value_col2 = row.iloc[1]
        
        # Check if " [" is present in the cell of the second column
        if " [" in cell_value_col2:
            # Remove " [" and all following characters from the second column
            row[1] = cell_value_col2.split(" [")[0]

        # Get the value in the third column "Entwicklung | Abnahme"
        cell_value = row.iloc[2]
        
        # Check if " | " is present in the cell
        if " | " in cell_value:
            # Split the cell value into multiple parts based on " | "
            parts = cell_value.split(" | ")
            
            # Iterate through each part and create a new row
            for part in parts:
                # Duplicate the row
                duplicated_row = row.copy()
                
                # Erase all characters after and including " | " in the third column
                duplicated_row[2] = part
                
                # Append the modified row to the output DataFrame
                output_df = pd.concat([output_df, duplicated_row.to_frame().transpose()], ignore_index=True)
        else:
            # If " | " is not present, include the original row
            output_df = pd.concat([output_df, row.to_frame().transpose()], ignore_index=True)
        
 
        # Save the output DataFrame to an Excel file
    output_df.to_excel(output_file, index=False)


if __name__ == "__main__":

# Specify file paths and column names
  excel_file2_path = 'u:\Temp\Programming\HLBDM\Data\AWSDB-Server.xlsx'
  OutputExcel1 = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2Stage-Step1.xlsx'
  OutputExcel2 = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2Stage-Intermediate.xlsx'

# Call the function
  formatSheet(excel_file2_path, OutputExcel1)
  multiplestages(OutputExcel1, OutputExcel2)
  
  


