import pandas as pd
import ipaddress

def is_valid_subnet(subnet):
    try:
        ipaddress.IPv4Network(subnet, strict=False)
        return True
    except ValueError:
        return False

def is_valid_ip(ip):
    try:
        ipaddress.IPv4Address(ip)
        return True
    except ValueError:
        return False

def add_subnet_column(sheet1_path, sheet2_path, subnet_column_name, ip_column_name, output_sheet_path):
    # Read the first Excel sheet
    df1 = pd.read_excel(sheet1_path)

    # Read the second Excel sheet
    df2 = pd.read_excel(sheet2_path)

    # Validate subnets in sheet 1
    valid_subnets = df1[df1[subnet_column_name].apply(is_valid_subnet)]

    # Validate IP addresses in sheet 2
    valid_ips = df2[df2[ip_column_name].apply(is_valid_ip)]

    # Function to find subnet for an IP
    def find_subnet(ip):
        for subnet in valid_subnets[subnet_column_name]:
            if ipaddress.IPv4Address(ip) in ipaddress.IPv4Network(subnet, strict=False):
                return subnet
        return None

    # Add a new "Subnet" column to sheet 2
    df2['Subnet'] = valid_ips[ip_column_name].apply(find_subnet)

    # Drop rows where cells in Subnet column
    df2 = df2.dropna(subset=[df2.columns[2]])

    # Write the updated sheet 2 to a new Excel file
    with pd.ExcelWriter(output_sheet_path, engine='xlsxwriter') as writer:
        df2.to_excel(writer, sheet_name='Sheet2_With_Subnets', index=False)

    return df2

# Usage
Subnetspath = 'u:\Temp\Programming\HLBDM\Dataoutput\SubNetz-FITS-S1.xlsx'
IPpath = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2Sub.xlsx'
subnet_column_name = 'Unnamed: 8'  # Replace with the actual column name
ip_column_name = 'IP Address'    # Replace with the actual column name
output_sheet_path = 'u:\Temp\Programming\HLBDM\Dataoutput\Rel-Srv2FITSSubn.xlsx'

df2_with_subnets = add_subnet_column(Subnetspath, IPpath, subnet_column_name, ip_column_name, output_sheet_path)

print("Sheet 2 with Subnets:")
print(df2_with_subnets)
