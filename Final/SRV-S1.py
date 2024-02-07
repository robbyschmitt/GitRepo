import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET




def step1(InputExcel, OutputExcel2):
    """Format the Rel-Srv2Stage-Intermediate Input sheet so we can work with it  """
    df = pd.read_excel(InputExcel, header=None)

    columns_to_drop = list([0, 2])
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Make a column for the Name description which is empty
    df.insert(1, "Name", df.iloc[0:, 0])

    # Add the suffix "-Node" to all values in the second column starting from the second row
    df.iloc[1:, 0] = df.iloc[1:, 0].astype(str) + '-Node'

   
    # Insert xsitype
    df.insert(1, "xsitype", range(1, len(df) + 1), True)
    df.iloc[:, 1] = "Device" 
    
    
    # Delete all duplicate Servers
    df = df.drop_duplicates(subset=df.columns[0], keep='first')

    # Convert all cells in the first column to uppercase so we can match the Servers in the CMDB-Datenbanken Instances file which have different cases.
    df.iloc[:, 0] = df.iloc[:, 0].str.upper()

     # Save the modified DataFrame to a new Excel file
    df.to_excel(OutputExcel2, index=False, header=False)


def open_excel_sheet(OutputExcel2):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(OutputExcel2)
    
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
    InputExcel = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2Stage-Intermediate.xlsx'
    OutputExcel2 = 'u:\Temp\Programming\HLBDM\Dataoutput\Server-S1-Elem.xlsx'
    xml_file_path = 'u:\Temp\Programming\HLBDM\Model\Imp-Server-Elem.xml'
   
  
    step1(InputExcel, OutputExcel2)
    convertxml()


    print(f"Processing completed. Modified data saved to {OutputExcel2}")