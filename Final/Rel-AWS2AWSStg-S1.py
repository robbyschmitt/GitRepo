import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

def modify_excel(input_file, output_file):
    # Read the existing Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Add a new column in the first position with the header ID
    df.insert(0, 'ID2', [f'Rel-AWS-AWSNode-{i:05}' for i in range(1, len(df) + 1)])

    # Add another new column in the second position
    df.insert(1, 'New_Column2', '')

    # Read contents of the cells in the now third column and copy the string before "_"
    df['New_Column2'] = df.iloc[:, 2].str.split('_').str[0]

    # Replace the XSI "Node" eith the Association Type needed
    df.iloc[:, 3] = "Composition"

    # Save the modified DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

def open_excel_sheet(file_path):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(file_path)
    
    # Assuming you want to work with the first sheet (index 0)
    sheet = workbook.worksheets[0]
    
    return workbook, sheet

def create_xml_from_excel(sheet):
    root = ET.Element("elements")

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


# Specify the input and output file names
input_excel_file = 'u:\Temp\Programming\HLBDM\Dataoutput\AWSStage-S4.xlsx'
output_excel_file = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-AWS2Stage.xlsx'
output_xml ='u:\Temp\Programming\HLBDM\Model\Rel-AWS2Stage.xml'

# Call the function to modify the Excel file
modify_excel(input_excel_file, output_excel_file)
xmlconv(output_excel_file, output_xml)
 
print(f"Excel sheet has been modified using pandas and saved as {output_excel_file}")
