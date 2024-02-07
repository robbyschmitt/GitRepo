import pandas as pd

def process_excel_file(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Remove the characters ".hlb.helaba.de" from the "Asset Key" column
    df['Asset Name'] = df['Asset Name'].str.replace('.hlb.helaba.de', '')

    # Drop rows where "Asset Name" is empty
    df = df.dropna(subset=['Asset Name'])
    df = df[~df['Asset Name'].str.startswith('DEFAULT')]

    # Convert all characters in the "Asset Key" column to uppercase
    df['Asset Name'] = df['Asset Name'].str.upper()

    # Define the valid values for the "Asset Type" column
    valid_asset_types = ['Domain controller', 'General', 'Server', 'Virtualization host', 'Specialized']

    # Drop rows where "Asset Type" is not in the valid list
    df = df[df['Asset Type'].isin(valid_asset_types)]

    # Drop all Defalut Types from the General selection - cant match these
    df = df[~df['Asset Name'].str.startswith('DEFAULT')]

    # Write the processed data to a new Excel file
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    input_file_path = "u:\Temp\Programming\HLBDM\Data\HLB-Assets.xlsx"  # Change this to your input file path
    output_file_path = "u:\Temp\Programming\HLBDM\Dataoutput\HLB-Assets1.xlsx"  # Change this to your desired output file path

    process_excel_file(input_file_path, output_file_path)
