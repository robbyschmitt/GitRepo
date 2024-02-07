import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

def remove_rows_with_square_brackets(input_file, output_file, sheet_name, column_name):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Find rows where the specified column contains square brackets
    rows_to_remove = df[df[column_name].str.contains('\[.*\]')]

    # Remove the rows with square brackets
    df = df[~df.index.isin(rows_to_remove.index)]

    # Save the updated DataFrame to a different Excel file
    df.to_excel(output_file, sheet_name=sheet_name, index=False)

def open_excel_sheet(output_file):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(output_file)
    
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
    # Specify the input file, output file, sheet name, and column name
    input_file = "u:\Temp\Programming\HLBDM\Dataoutput\Rel-NodeStage-Step2.xlsx"
    output_file = "u:\Temp\Programming\HLBDM\Dataoutput\Rel-NodeStage-Step3.xlsx"
    xml_file_path = "u:\Temp\Programming\HLBDM\Model\Imp-Srv2AWSStage-Rel.xml"
    sheet_name = "Sheet1"
    column_name = "Server"

    # Call the function to remove rows with square brackets and save the results
    remove_rows_with_square_brackets(input_file, output_file, sheet_name, column_name)
    try:
        workbook, sheet = open_excel_sheet(output_file)
        
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
