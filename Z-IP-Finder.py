
####################

import pandas as pd
import ipaddress

# Function to check if a value is a valid subnet
def is_valid_subnet(subnet):
    try:
        ipaddress.IPv4Network(subnet)
        return True
    except ValueError:
        return False

def is_valid_ip(ip):
    try:
        ipaddress.IPv4Address(ip)
        return True
    except ipaddress.AddressValueError:
        return False


# Function to find the matching subnet for an IP address
def find_matching_subnet(ip_address, subnets):
    ip = ipaddress.IPv4Address(ip_address)
    for subnet in subnets:
        if ip in subnet:
            return subnet
    return None
output_df = pd.concat([output_df, duplicated_row.to_frame().transpose()], ignore_index=True)
 

# Read the first Excel sheet containing network subnets
subnet_sheet = pd.read_excel('u:\Temp\Programming\HLBDM\Dataoutput\SubNetz-FITS-S1.xlsx', header=None, names=['Unnamed: 8'])

# Create a new DataFrame for invalid subnets
invalid_subnets = subnet_sheet[~subnet_sheet['Unnamed: 8'].apply(is_valid_subnet)]

# Save the DataFrame with invalid subnets to a new Excel sheet
invalid_subnets.to_excel('u:\Temp\Programming\HLBDM\Dataoutput\ACH-Bad-Netzwerksegmente-FITS.xlsx', index=False)

# Remove rows with invalid subnet values
subnet_sheet = subnet_sheet[subnet_sheet['Unnamed: 8'].apply(is_valid_subnet)]

# Convert valid subnets to IPv4Network objects
subnets = [ipaddress.IPv4Network(subnet) for subnet in subnet_sheet['Unnamed: 8']]

# Read the second Excel sheet containing IP addresses
ip_sheet = pd.read_excel('u:\Temp\Programming\HLBDM\Dataoutput\zzz.xlsx', header=None)

if is_valid_ip(ip_sheet[1]):
# Append the modified row to the output DataFrame
  ip_sheet = pd.concat([ip_sheet, ip_sheet.to_frame().transpose()], ignore_index=True)
 

# Find and write the appropriate subnet for each IP address
ip_sheet.iloc[:, 2] = ip_sheet.iloc[:, 1].apply(lambda ip: find_matching_subnet(ip, subnets))

# Save the updated second Excel sheet
ip_sheet.to_excel('u:\Temp\Programming\HLBDM\Dataoutput\IP-ServerSubnet.xlsx', index=False)

