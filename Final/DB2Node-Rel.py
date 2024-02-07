import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

def DBprep(input_excel_file_path, DBprep_path):
    # Read the Excel file
    df = pd.read_excel(input_excel_file_path, engine='openpyxl', skiprows=0)

    # Remove Duplicates in ID column 0
    df = df.drop_duplicates(subset=df.columns[1], keep='first')


   # Insert new column with index
    df.insert(0, "NewColumn", range(1, len(df) + 1), True)
    df.iloc[:, 0] = "Rel-DB2Stage-" + df.iloc[:, 0].astype(str).str.zfill(5) 

   # Insert new column with index
    df.insert(1, "XSI", "Association")
    
    # Make the source stage
    df.iloc[:, 6] =  df.iloc[:, 6] + '_' + df.iloc[:, 8] 
 
    # Drop colums not needed for element xsi
    columns_to_drop = [2, 4, 5, 7, 8, 9]
    df = df.drop(df.columns[columns_to_drop], axis=1)
   
   # Insert new empty Name and Desc columns 
    df.insert(4, "Name", "")
    df.insert(5, "Descr", "")
  

    # Save the modified DataFrame to a new Excel file
    df.to_excel(DBprep_path, index=False)




def process_excel(input_file, output_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Add values from column 5 to the cell value in column 7 with "_" as a separator
    df[7] = df.apply(lambda row: f"{row[4]}_{row[6]}", axis=1)

    # Save the results to a new sheet in the same Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Result', index=False)

if __name__ == "__main__":
    # Provide the input and output file paths
    input_excel_path = 'u:\Temp\Programming\HLBDM\Dataoutput\DB-Main.xlsx'
    DBprep_path = 'u:\Temp\Programming\HLBDM\Dataoutput\DB-Prep.xlsx'
    output_excel_path = 'u:\Temp\Programming\HLBDM\Dataoutput\DBInst2Node-Rel.xlsx'

    # Process the Excel file
    #process_excel(input_excel_path, output_excel_path)
    DBprep(input_excel_path, DBprep_path)
    print(f"Processing complete. Results saved to {output_excel_path}")
