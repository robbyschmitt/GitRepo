import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET

# First we import and configure the AWS File
def process_AWSexcel_file(input_AWSfile, output_AWSfile):
    # Read the Excel file into a pandas DataFrame
    rowprod = pd.read_excel(input_AWSfile)

    # Copy each row four times for the different stages which we dont know
    rowabn = rowprod.copy()
    rowtst = rowprod.copy()
    rowent = rowprod.copy()
    
    # Apply the logic of the ID
    rowprod['Allgemein'] = rowprod['Allgemein'] + '_Produktion' 
    rowabn['Allgemein'] = rowabn['Allgemein'] + '_Abnahme' 
    rowtst['Allgemein'] = rowtst['Allgemein'] + '_Test' 
    rowent['Allgemein'] = rowent['Allgemein'] + '_Entwicklung' 

    # Apply the logic of the Name
    rowprod['Unnamed: 1'] = rowprod['Unnamed: 1'] + '_Produktion' 
    rowabn['Unnamed: 1'] = rowabn['Unnamed: 1'] + '_Abnahme' 
    rowtst['Unnamed: 1'] = rowtst['Unnamed: 1'] + '_Test' 
    rowent['Unnamed: 1'] = rowent['Unnamed: 1']+ '_Entwicklung' 

    # String them together as a whole called df 
    df = pd.concat([rowprod, rowabn, rowtst, rowent], ignore_index=True) 

    # Delete any Blanks in the ID-Column as well as bad characters
    df['Allgemein'] = df['Allgemein'].apply(lambda x: '_'.join(x.split()) if isinstance(x, str) else x)

    # Set the xsi to Node
    df['NewColumn'] = 'ApplicationComponent'

    # Save the modified DataFrame to a new Excel file
    df.to_excel(output_AWSfile, index=False)
   
def open_excel_sheet(input_file_path):
    # Open the Excel workbook
    workbook = openpyxl.load_workbook(input_file_path)
    
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

def convert():
    try:
        workbook, sheet = open_excel_sheet(output_file_path2)
        
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
    # Replace these with your file paths and column name
    input_file_path = "u:\Temp\Programming\HLBDM\Data\ACH-AWS-S4.xlsx"
    xml_file_path = 'u:\Temp\Programming\HLBDM\Model\Imp-AWSStage-Elem.xml'
    output_file_path2 = "u:\Temp\Programming\HLBDM\Dataoutput\AWSStage-S4.xlsx"
    sheet_name = "Sheet1"  # Change this to your sheet name

    process_AWSexcel_file(input_file_path, output_file_path2)
    convert()
