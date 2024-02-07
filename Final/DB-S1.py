import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

def DBMain(input_excel_file_path, DBMain_path):
    # Read the Excel file
    df = pd.read_excel(input_excel_file_path, engine='openpyxl', skiprows=0)

    # Copy column three and insert as a new column "Descr"
    df.insert(2, 'Descr', df.iloc[:, 3])

    # Drop rows where the value of the first column is not "Bereitgestellt"
    df = df[df.iloc[:, 0] == 'Bereitgestellt']

    # Replace " auf " with "-"
    df['Descr'] = df['Descr'].str.replace(' auf ', '-')

    # Replace "\" with "-"
    df['Descr'] = df['Descr'].str.replace('\\', '-')
    df['Descr'] = df['Descr'].replace(to_replace={' ': '_'}, regex=True)

    # Replace characters in front of "auf " with "DBInst-"
    df.iloc[:, 3] = df.iloc[:, 3].str.replace('.*auf ', 'DBInst-', regex=True)
    df.iloc[:, 3] = df.iloc[:, 3].str.replace('\\', '-')

    # Drop columns 1, 2, 7, 8, 9, 11, 12, 13, 18, 19, 20, 21, 22
    columns_to_drop = [0, 1, 6, 7, 8, 9, 10, 11, 12, 17, 18, 19, 20, 21]
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Change the case of cells in column 2 to uppercase
    df.iloc[:, 1] = df.iloc[:, 1].str.upper()

    # Connect DB to Stage"
    df.iloc[:, 4] = df.iloc[:, 4] + '_' + df.iloc[:, 6]
    

    # Save the modified DataFrame to a new Excel file
    df.to_excel(DBMain_path, index=False)

def DBElements(DBMain_path, DBElem):
    # Read the Excel file
    df = pd.read_excel(DBMain_path, engine='openpyxl', skiprows=0)
    
    # Insert XSI
    df.insert(1, 'XSI', 'TechnologyService') 
    
    # Drop colums not needed for element xsi
    columns_to_drop = [2, 4, 5, 7, 8]
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Remove Duplicates in ID column 0
    df = df.drop_duplicates(subset=df.columns[0], keep='first')

    # Save the modified DataFrame to a new Excel file
    df.to_excel(DBElem, index=False)

def open_excel_sheet(file_path):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(file_path)
    
    # Assuming you want to work with the first sheet (index 0)
    sheet = workbook.worksheets[0]
    
    return workbook, sheet

def create_xml_from_excel(sheet):
    root = ET.Element("elements")

    # Iterate through rows in the sheet
    for row in sheet.iter_rows(min_row=2, values_only=True):
        element_identifier, xsi_type, name_de, documentation = row

        element = ET.SubElement(root, f"element", attrib={'identifier': element_identifier, 'xsi:type': xsi_type})

        name = ET.SubElement(element, f"name", attrib={'xml:lang': 'en'})
        name.text = name_de

        doc = ET.SubElement(element, f"documentation")
        doc.text = documentation


    return root

def save_xml(xml_root, xml_file_path):
    tree = ET.ElementTree(xml_root)
    tree.write(xml_file_path, encoding="utf-8", xml_declaration=True)
    print(f"XML file saved to {xml_file_path} successfully.")

def xmlconv(file_path,xml_file_path):
    try:
        workbook, sheet = open_excel_sheet(file_path)
        
        # Create XML from the Excel sheet
        xml_root = create_xml_from_excel(sheet)
        
        # Save XML to a file
        save_xml(xml_root, xml_file_path)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Close the workbook
        if 'workbook' in locals():
            workbook.close()

if __name__ == "__main__":
    # Define file paths
    input_excel_file_path = r"u:\Temp\Programming\HLBDM\Data\CMDB - Datenbanken mit AWS.xlsx"
    DBMain_path = r"u:\Temp\Programming\HLBDM\Dataoutput\DB-S1.xlsx"
    DBElements_path = r"u:\Temp\Programming\HLBDM\Dataoutput\DB-S2.xlsx"
    output_xml = r"u:\Temp\Programming\HLBDM\Model\Imp-DB-Elem.xml"

    # Call the main function
    DBMain(input_excel_file_path, DBMain_path)
    DBElements(DBMain_path, DBElements_path)
    xmlconv(DBElements_path, output_xml)