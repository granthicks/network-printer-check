import csv
import subprocess
import sys
import pandas as pd
from tqdm import tqdm
import fnmatch

"""
FUNCTIONS
"""
# Importing list of hostnames
def import_hostnames_list(file_name):
    """
    Returns a list of hostnames imported from a csv file
    """
    with open(file_name) as f:
        reader = csv.reader(f)
        data = list(reader)
    return [x[0] for x in data]

# Powershell script to find connected printers
def pshell_printer_script(hostname):
    """
    Returns a dictionary of printer names and ip addresses associated with
    the given hostname
    """

    # Printer names to ignore
    name_ignore_terms = [
    'Send To OneNote 2013', 
    'Send to Microsoft', 
    'Microsoft XPS Document Writer', 
    'Microsoft Print to PDF', 
    'Fax',
    '',
    'Adobe PDF'
    ]

    # Printer ports to ignore
    port_ignore_terms = [
        'nul:',
        'PORTPROMPT:',
        'SHRFAX:',
        '',
        'Documents"\\"*.pdf'
    ]

    # Powershell command for grabbing printer port info for hostname
    pn_cmd = f"get-printer -computername {hostname} | Select-Object portname"
    pn_completed = subprocess.run(["powershell","-Command", pn_cmd], capture_output=True, text=True)
    
    # Powershell command for grabbing printer name information for hostname
    n_cmd = f"get-printer -computername {hostname} | Select-Object name"
    n_completed = subprocess.run(["powershell","-Command", n_cmd], capture_output=True, text=True)
    
    # Formatting the results for ports/names and filtering out unneeded information
    pn_result = pn_completed.stdout.splitlines()[3:]
    pn_result = [x.rstrip() for x in pn_result]
    pn_result = [x for x in pn_result if x not in port_ignore_terms]
    n_result = n_completed.stdout.splitlines()[3:]
    n_result = [n.rstrip() for n in n_result]
    n_result = [n for n in n_result if n not in name_ignore_terms]

    # Saving to dictionary
    printer_result = [{'printer_name':n, 'printer_ip':pn} for n,pn in zip(n_result,pn_result)]
    if not printer_result:
        printer_result = {"printer_name":"NONE", "printer_ip":"NONE"}
    return printer_result

def create_dataframe():
    """
    Creates dataframe to be updated with printer information and exported to csv
    """
    return pd.DataFrame(columns=['hostname', 'printer_name', 'printer_ip'])    

def export_csv(data, final_name='network-printers-checked.csv'):
    """
    Exports data as new csv with given filename
    """
    data.to_csv(final_name, index=False)

def run(file_name):
    hostnames = import_hostnames_list(file_name)

    result_dict = {}
    print("Collecting printer information.")
    for hostname in tqdm(hostnames):
        result_dict[hostname] = pshell_printer_script(hostname)
    
    export_df = create_dataframe()

    print("Converting data.")
    for hostname in tqdm(result_dict.keys()):
        for printer in result_dict[hostname]:
            try:
                temp_dict = {
                    "hostname":hostname,
                    "printer_name":printer['printer_name'],
                    "printer_ip":printer['printer_ip']
                }
            except:
                temp_dict = {
                    "hostname":hostname,
                    "printer_name":"NONE",
                    "printer_ip":"NONE"
                }
            export_df = export_df.append(temp_dict, ignore_index=True)
    return export_df

"""
MAIN PROGRAM
"""

if __name__ =='__main__':
    pattern = "*.csv"
    file_to_run = sys.argv[0]

    if not fnmatch.fnmatch(file_to_run, pattern):
        while not fnmatch.fnmatch(file_to_run, pattern):
            file_to_run = input("Please enter the path to the csv file to use (or 'exit' to quit): ")

    if file_to_run.lower == 'exit':
        print("Goodbye.")
        sys.exit()
    else:
        to_export = run(file_to_run)
        saved = False
        while not saved:
            try:
                export_csv(to_export)
                saved = True
            except:
                print("Unable to save file with current filename")
                export_name = input("Please enter a name for export file: ") + ".csv"
                export_csv(to_export, export_name)
                saved = True
