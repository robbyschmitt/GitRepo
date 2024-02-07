import ipaddress
import pandas as pd

def find_subnet(ip, network_blocks):
    for network_block in network_blocks:
        network = ipaddress.IPv4Network(network_block, strict=False)
        if ipaddress.IPv4Address(ip) in network:
            return network_block
    return "Not Found"

def main():
    # Replace these file paths with your actual Excel files
    excel_file_path_1 = 'u:\Temp\Programming\HLBDM\Dataoutput\ACH-Netzwerksegmente-FITS.xlsx'
    excel_file_path_2 = 'u:\Temp\Programming\HLBDM\Dataoutput\IP-Server.xlsx'

    # Read the network address blocks from Excel sheet 1
    df_network_blocks = pd.read_excel(excel_file_path_1)
    network_blocks = df_network_blocks['NetSeg_ID'].tolist()

    # Read the IP addresses from Excel sheet 2
    df_ip_addresses = pd.read_excel(excel_file_path_2)
    ip_addresses = df_ip_addresses['Unnamed: 1'].tolist()

    # Find the appropriate subnet for each IP address
    results = []
    for ip in ip_addresses:
        subnet = find_subnet(ip, network_blocks)
        results.append({'Unnamed: 1': ip, 'NetSeg_ID': subnet})

    # Create a new DataFrame with the results
    df_results = pd.DataFrame(results)

    # Write the results to a new Excel file
    output_excel_file = 'u:\Temp\Programming\HLBDM\Dataoutput\output_results.xlsx'
    df_results.to_excel(output_excel_file, index=False, engine='openpyxl')

    print(f"Results written to {output_excel_file}")

if __name__ == "__main__":
    main()
