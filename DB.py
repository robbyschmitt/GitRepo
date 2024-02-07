import pandas as pd

def DBInst_S1(input_file, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file, engine='openpyxl', skiprows=3)

    # Drop rows where the fifth column doesn't have the value "Bereitgestellt"
    df = df[df.iloc[:, 4] == "Bereitgestellt"]

    # Drop rows where the fourteenth column matches the value "T-Systems"
    df = df[df.iloc[:, 13] != "T-Systems"]

    # Drop columns 1 to 7 and 11 to 26
    df = df.drop(df.columns[list(range(0, 7)) + list(range(10, 26))], axis=1)
    
    # Copy the second column as a new column at position 1
    df.insert(1, 'ID', df.iloc[:, 1])

    # Insert the prefix "DBInst-" in front of the values in column 2
    df.iloc[:, 1] = "DBInst-" + df.iloc[:, 1].astype(str)

    # Replace all strings in column 2 with the value "/" with "-"
    df.iloc[:, 1] = df.iloc[:, 1].str.replace("\\", "-")

    # Insert new column at position 3
    df.insert(3, 'Description', df.iloc[:, 0])

    # Drop the first column which was duplicated to column 4
    df = df.iloc[:, 1:]

    # Change the case of cells in column 2 to uppercase
    df.iloc[:, 0] = df.iloc[:, 0].str.upper()


    # Save the results to a new Excel file
    df.to_excel(output_file, index=False)


if __name__ == "__main__":
     input_file = r"u:\Temp\Programming\HLBDM\Data\CMDB - Datenbankmanagementsysteme.xlsx"
     output_file = r"u:\Temp\Programming\HLBDM\Dataoutput\DBInst-S1.xlsx" 

     DBInst_S1(input_file, output_file)
     print(f"Saved resilts to {output_file}")
