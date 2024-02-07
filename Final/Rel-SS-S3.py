import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET


def process_excel(input_file, output_excel_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Make the ID unique
    df.iloc[:, 0] = "Rel-AWS2SS-N-" + (df.index + 1).map(lambda x: f"{x:05d}")
    
    # new fifth column
    df.insert(5, 'New', '')

    # Change all values in the fourth column to "Association"
    df['New'] = df['Unnamed: 8']


    # Drop the sixth column
    df = df.drop(columns=['Unnamed: 8'])
    df = df.drop(columns=['Unnamed: 13'])

    # Save the results to a new Excel file
    df.to_excel(output_excel_file, index=False)

def open_excel_sheet(output_excel_file):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(output_excel_file)
    
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

if __name__ == "__main__":
    input_excel_file = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SSFin-Rel.xlsx"  # Change this to your input Excel file
    output_excel_file = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SSFinal-Rel.xlsx"  # Change this to your desired output Excel file
    xml_file_path = "u:\Temp\Programming\HLBDM\Dataoutput\Imp-SSFinal-Rel.xml"


    process_excel(input_excel_file, output_excel_file)
    try:
        workbook, sheet = open_excel_sheet(output_excel_file)
        
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
