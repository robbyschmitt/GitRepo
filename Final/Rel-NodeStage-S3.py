import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

def add_suffix_to_column(file_path, output_file_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, header=None)

    # Add the suffix "-Node" to all values in the second column starting from the second row to get source
    df.iloc[1:, 1] = df.iloc[1:, 1].astype(str) + '-Node'

    # Convert all cells in the first column to uppercase so we can match the Servers in the other Instances file which have different cases.
    df.iloc[:, 1] = df.iloc[:, 1].str.upper()

    # Add the values IS to Umgebung to create Traget
    df.iloc[1:, 0] = df.iloc[1:, 0].astype(str) + '_' + df.iloc[1:, 2].astype(str)

   # Insert new column 
    df.insert(0, "ID", range(1, len(df) + 1), True)
    df.iloc[:, 0] = "Rel-Srv2Stage-N-" + df.iloc[:, 0].astype(str).str.zfill(5) 

  
    # Insert xsitype
    df.insert(3, "xsitype", range(1, len(df) + 1), True)
    df.iloc[:, 3] = "Association" 
    
    # Insert Name
    df.insert(4, "Name", range(1, len(df) + 1), True)
    df.iloc[:, 4] = "Relation" 

    # Delete the fifth column
    columns_to_drop = list([5])
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Save the modified DataFrame to a new Excel file
    df.to_excel(output_file_path, index=False, header=False)

def open_excel_sheet(OutputExcel2):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(OutputExcel2)
    
    # Assuming you want to work with the first sheet (index 0)
    sheet = workbook.worksheets[0]
    
    return workbook, sheet

def create_xml_from_excel(sheet):
    # Iterate through rows in the sheet
    root = ET.Element("relationships")

    # Iterate through rows in the sheet
    for row in sheet.iter_rows(min_row=2, values_only=True):
        element_identifier, source, target, xsi_type, name_en, documentation = row

        element = ET.SubElement(root, f"relationship", attrib={'identifier': element_identifier, 'source': source, 'target': target, 'xsi:type': xsi_type})
        name = ET.SubElement(element, f"name", attrib={'xml:lang': 'en'})
        name.text = name_en
        
        doc = ET.SubElement(element, f"documentation")
        doc.text = documentation
        
    return root


    return root

def save_xml(xml_root, xml_file_path):
    tree = ET.ElementTree(xml_root)
    tree.write(xml_file_path, encoding="utf-8", xml_declaration=True)
    print(f"XML file saved to {xml_file_path} successfully.")

def convertxml():
    try:
        workbook, sheet = open_excel_sheet(output_file_path)
        
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
    input_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2Stage-Intermediate.xlsx'
    output_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-NodeStage-Step2.xlsx'
    xml_file_path = 'u:\Temp\Programming\HLBDM\Model\Imp-Srv2Stage.xml'

    add_suffix_to_column(input_file_path, output_file_path)
    convertxml()

    print(f"Processing completed. Modified data saved to {output_file_path}")