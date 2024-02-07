import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

def step1(InputExcel, OutputExcel2):
    """Format the Rel-Srv2Stage-Intermediate Input sheet so we can work with it  """
    df = pd.read_excel(InputExcel, header=None)
   
   # Insert new column and number it for the ID
    df.insert(0, "ID", range(1, len(df) + 1))
    df.iloc[:, 0] = "Rel-Srv2Subnet-N-" + df.iloc[:, 0].astype(str).str.zfill(5) 

   # Add the suffix "-Node" to all Servernames column starting from the second row
    df.iloc[1:, 1] = df.iloc[1:, 1].astype(str) + '-NODE'

    # Insert new column and number it for the ID
    df.insert(3, "Name", range(1, len(df) + 1), True)
    # Copy the IP address
    df.iloc[1:, 3] = df.iloc[1:, 2]
    # Copy the Subnet as target
    df.iloc[1:, 2] = df.iloc[1:, 4]
    # Rename IP Adresses 
    df.iloc[1:, 2] = 'IP-' + df.iloc[1:, 2].replace(to_replace={'/': '-'}, regex=True) + '-FITS'
    # Specify the XSI as Association
    df.insert(3, "XSI", df.iloc[0:, 4], True)
    df.iloc[1:, 3] = 'Association'

     # Save the modified DataFrame to a new Excel file
    df.to_excel(OutputExcel2, index=False, header=False)

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

def save_xml(xml_root, xml_file_path):
    tree = ET.ElementTree(xml_root)
    tree.write(xml_file_path, encoding="utf-8", xml_declaration=True)
    print(f"XML file saved to {xml_file_path} successfully.")

def convertxml():
    try:
        workbook, sheet = open_excel_sheet(OutputExcel2)
        
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
    # Varaibles
    InputExcel = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2FITSSubn.xlsx'
    OutputExcel2 = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2FITSSubn-S2.xlsx'
    xml_file_path = 'u:\Temp\Programming\HLBDM\Model\Imp-Srv2Subnet-Rel.xml'

    step1(InputExcel, OutputExcel2)
    convertxml()

    print(f"Processing completed. Modified data saved to {convertxml}") 
