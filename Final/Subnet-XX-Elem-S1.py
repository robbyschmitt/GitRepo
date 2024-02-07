import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET


def process_fits(input_file, output_fits):
    # Read the Excel sheet
    df = pd.read_excel(input_file)

    # Drop the first three rows
    df = df.iloc[3:]

    # Filter rows based on the fourteenth column
    df = df[df.iloc[:, 13] == "FI-TS"]

    # Filter rows based on the fifth column
    df = df[df.iloc[:, 4] == "Bereitgestellt"]

    # Drop specified columns
    df = df.drop(df.columns[[0, 1, 2, 3, 4, 5, 6, 9, 11, 12, 15, 16, 17, 18, 19, 20, 21, 22]], axis=1)

    # Copy column 2 (Subnet) as a new column in the first position
    df.insert(0, 'New_Column', df.iloc[:, 1])


    # Ad a Character for Archi ID - Archi doesnt accept IDs starting with a number. Then Replace harmful archi index characters in the ID column
    df.iloc[:, 0] = "IP-" + df.iloc[:, 0]
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("/", "-")
 
    # Ad FITS as Suffix to identify Subnetprov
    df.iloc[:, 1] = df.iloc[:, 1] + "-FITS"
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("/", "-")


    # Add a suffix to the first column based on the fourteenth column
    df.iloc[:, 0] = df.iloc[:, 0] + "-" + df.iloc[:, 4].astype(str)
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("FI-TS", "FITS")

    # Insert xsitype in second column
    df.insert(1, "xsitype", range(1, len(df) + 1), True)
    df.iloc[:, 1] = "CommunicationNetwork" 
  
    # Drop the rest
    columns_to_drop = list([4, 5, 6])
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Remove Blanks
    df['New_Column'] = df['New_Column'].apply(lambda x: '_'.join(x.split()) if isinstance(x, str) else x)

     # Save the processed data to a new Excel file
    df.to_excel(output_tsy, index=False)

     # Save the processed data to a new Excel file
    df.to_excel(output_fits, index=False)

def process_tsy(input_file, output_tsy):
    # Read the Excel sheet
    df = pd.read_excel(input_file)

    # Drop the first three rows
    df = df.iloc[3:]

    # Filter rows based on the fourteenth column
    df = df[df.iloc[:, 13] == "T-Systems"]

    # Filter rows based on the fifth column
    df = df[df.iloc[:, 4] == "Bereitgestellt"]

    # Drop specified columns
    df = df.drop(df.columns[[0, 1, 2, 3, 4, 5, 6, 9, 11, 12, 15, 16, 17, 18, 19, 20, 21, 22]], axis=1)

    # Copy column 2 (Subnet) as a new column in the first position
    df.insert(0, 'New_Column', df.iloc[:, 1])


    # Ad a Character for Archi ID - Archi doesnt accept IDs starting with a number. Then Replace harmful archi index characters in the ID column
    df.iloc[:, 0] = "IP-" + df.iloc[:, 0]
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("/", "-")

    # Ad TSY as Suffix to identify Subnetprov
    df.iloc[:, 1] = df.iloc[:, 1] + "-TSY"
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("/", "-")


    # Add a suffix to the first column 
    df.iloc[:, 0] = df.iloc[:, 0] + "-" + df.iloc[:, 4].astype(str)

    # Insert xsitype in second column
    df.insert(1, "xsitype", range(1, len(df) + 1), True)
    df.iloc[:, 1] = "CommunicationNetwork" 
  
    # Drop the rest
    columns_to_drop = list([4, 5, 6])
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Remove Blanks
    df['New_Column'] = df['New_Column'].apply(lambda x: '_'.join(x.split()) if isinstance(x, str) else x)
     # Save the processed data to a new Excel file
    df.to_excel(output_tsy, index=False)

def process_icf(input_file, output_icf):
    # Read the Excel sheet
    df = pd.read_excel(input_file)

    # Drop the first three rows
    df = df.iloc[3:]

    # Filter rows based on the fourteenth column
    df = df[df.iloc[:, 13] == "ICF"]

    # Filter rows based on the fifth column
    df = df[df.iloc[:, 4] == "Bereitgestellt"]

    # Drop specified columns
    df = df.drop(df.columns[[0, 1, 2, 3, 4, 5, 6, 9, 11, 12, 15, 16, 17, 18, 19, 20, 21, 22]], axis=1)

    # Copy column 2 (Subnet) as a new column in the first position
    df.insert(0, 'New_Column', df.iloc[:, 1])


    # Ad a Character for Archi ID - Archi doesnt accept IDs starting with a number. Then Replace harmful archi index characters in the ID column
    df.iloc[:, 0] = "IP-" + df.iloc[:, 0]
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("/", "-")

    # Ad ICF as Suffix to identify Subnetprov
    df.iloc[:, 1] = df.iloc[:, 1] + "-ICF"
    df.iloc[:, 0] = df.iloc[:, 0].str.replace("/", "-")


    # Add a suffix to the first column based on the fourteenth column
    df.iloc[:, 0] = df.iloc[:, 0] + "-" + df.iloc[:, 4].astype(str)

    # Insert xsitype in second column
    df.insert(1, "xsitype", range(1, len(df) + 1), True)
    df.iloc[:, 1] = "CommunicationNetwork" 
  
    # Drop the rest
    columns_to_drop = list([4, 5, 6])
    df = df.drop(df.columns[columns_to_drop], axis=1)

    # Remove Blanks
    df['New_Column'] = df['New_Column'].apply(lambda x: '_'.join(x.split()) if isinstance(x, str) else x)

     # Save the processed data to a new Excel file
    df.to_excel(output_tsy, index=False)


     # Save the processed data to a new Excel file
    df.to_excel(output_icf, index=False)

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
    input_file_path = 'u:\Temp\Programming\HLBDM\Data\CMDB - Netzwerksegmente.xlsx'
    output_fits = r'U:\Temp\Programming\HLBDM\Dataoutput\SubNetz-FITS-S1.xlsx'
    output_tsy = r'U:\Temp\Programming\HLBDM\Dataoutput\SubNetz-TSY-S1.xlsx'
    output_icf = r'U:\Temp\Programming\HLBDM\Dataoutput\SubNetz-ICF-S1.xlsx'
    output_fitsxml = r'u:\Temp\Programming\HLBDM\Model\Imp-SubNetz-FITS-Elem.xml'
    output_tsyxml = r'u:\Temp\Programming\HLBDM\Model\Imp-SubNetz-TSY-Elem.xml'
    output_icfxml = r'u:\Temp\Programming\HLBDM\Model\Imp-SubNetz-ICF-Elem.xml'

    process_fits(input_file_path, output_fits)
    process_tsy(input_file_path, output_tsy)
    process_icf(input_file_path, output_icf)
    xmlconv(output_fits, output_fitsxml)
    xmlconv(output_tsy, output_tsyxml)
    xmlconv(output_icf, output_icfxml)
