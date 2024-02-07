# Manipulates the base export AWS-DB
#ACH 2023.12
# export column A, "Nr"; Column B, "Kurzname" und Column D "Beschreibung" from orig AWS-DB  and import as columns A,C und D in new Excel 
# Spalte B im neuen Excel mit dem Wert "ApplicationComponent" füllen.
# Dieses Python Script Teil1 auf Excel laufen lassen. 


import openpyxl
import xml.etree.ElementTree as ET

def open_excel_sheet(file_path):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(file_path)
    
    # Assuming you want to work with the first sheet (index 0)
    sheet = workbook.worksheets[0]
    
    return workbook, sheet

def create_xml_from_excel(sheet):
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

def main():
    # Specify the path to your Excel file
    file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\AWS-Imp-Step4.xlsx'
    
    # Specify the path for the XML file
    xml_file_path = 'u:\Temp\Programming\HLBDM\Dataoutput\Step5.xml'

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
    main()