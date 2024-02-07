import pandas as pd

# Specify the path to your Excel file
excel_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\Imp-AWSStage-Elem.xlsx'

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_file_path, engine='openpyxl')

# Replace blanks within strings in column 1 with "_"
df['AWS'] = df['AWS'].apply(lambda x: '_'.join(x.split()) if isinstance(x, str) else x)

# Save the modified DataFrame back to the Excel file
df.to_excel(excel_file_path, index=False, engine='openpyxl')