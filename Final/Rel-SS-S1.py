# Takes the AWS (DV-Anwendungen und TIA) File and sets it up for conversion.

import pandas as pd

def process_excel_file(input_file, output_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Delete the first three rows made by CMDB
    df = df.iloc[3:]

    # Delete specified columns (first, third, sixth, seventh, and all the rest)
    columns_to_drop = list(range(0, 6))+ [10, 12, 14] + list(range(16, 25))
    df = df.drop(df.columns[columns_to_drop], axis=1)

       # Filter nach Bereitgestellten Assets    
    # Filter column AssetLifecycleStage has the value "Bereitgestellt"
    df = df[df.iloc[:, 0] == 'Bereitgestellt']
 
    # ISxxx Produktion in IS Nummer convertieren
    df.iloc[:, 3] = df.iloc[:, 3].str.replace(r'\s*Produktion$', '', regex=True)
    df.iloc[:, 4] = df.iloc[:, 4].str.replace(r'\s*Produktion$', '', regex=True)
 
    # Delete Bereitgestellt columns 
    columns2_to_drop = list(range(1))
    df = df.drop(df.columns[columns2_to_drop], axis=1)

    # Check if the value  is "A<-B" and swap contents of corresponding cells AWSA and AWSB
    mask = df.iloc[:, 4] == 'A<-B'
    df.loc[mask, [df.columns[2], df.columns[3]]] = df.loc[mask, [df.columns[3], df.columns[2]]].values


    # Löscht alle ISext SS aus Ziel da nicht im Modell
    mask_delete = df.iloc[:, 3].str.contains('ext', case=False, na=False)
    df = df[~mask_delete]
    # Löscht alle ISext SS aus Target da nicht im Modell
    mask_delete = df.iloc[:, 2].str.contains('ext', case=False, na=False)
    df = df[~mask_delete]


    # Add a new column with cells containing the  "Flow" direction as the second column
    df.insert(4, 'NewColumn', 'Flow')


  # Save the filtered DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

# -------------------------------------------------------------

if __name__ == "__main__":
    # setup the file names
    input_file_name = 'u:\Temp\Programming\HLBDM\Data\CMDB - AWS Schnittstellen.xlsx'
    output_file_name = 'u:\Temp\Programming\HLBDM\Dataoutput\SchnittSt-S1.xlsx'
    
    process_excel_file(input_file_name, output_file_name)

    print(f"Excel file processed and saved as {output_file_name}")