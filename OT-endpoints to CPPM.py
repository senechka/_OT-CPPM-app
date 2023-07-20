"""
Program to prepare Damfoss OT Devices to import the CPPM as endpoints.
"""
# Imports
import pandas as pd
import re
from pathlib import Path
# Import tkinter to open file
import tkinter as tk
from tkinter import filedialog 
def openotfile():
    """
    Function to choose and open local (downloaded) OT file.
    Returns:
        path.name - path to file (incl.file name)
    """   
    #Choose OT file
    root = tk.Tk()
    root.withdraw
    print (Path.cwd())
    path = filedialog.askopenfile(initialdir=Path() / 'workdir/', title="Select file",
                    filetypes=(("XLS files", "*.xlsx"),("all files", "*.*")))
    return (path.name)
#Read OT sheet from file (if there is no tkinter installed, please change function call openotfile() to full path name)
df=pd.read_excel(openotfile(), sheet_name='OT Endpoint MAC Addresses',index_col=False)
df.columns = df.iloc[3]
df2 = df.tail(-4)
# Choose only rows started with "OT-""
df2 = df2.loc[df2["Aruba User Role"].str.startswith('OT-')]
#Choose only MAC and Roles columns and drop lines with empty MACs
df2 = df2[["Mac-Addresses","Aruba User Role"]].dropna(subset=["Mac-Addresses"])
#Normalise MACs for CPPM (delete : - . and space simbols) 
df2["Mac-Addresses"] = df2.loc[:,"Mac-Addresses"].str.replace('[:-]','', regex =True).replace(' ','', regex =True).replace('\.','', regex =True) # FIXME too long ?
#read template file
with open(Path() / 'templates/Endpoint-clear_templ.xml',"r") as xml_template:
    template_lines=xml_template.readlines()
xml_template.close
#write first part of template to OT XML file
outputxmlfile = Path() / 'workdir/OT-CPPM.xml'
with open(outputxmlfile,"w") as otxmlfile:
    otxmlfile.writelines(template_lines[0:4])
with open(outputxmlfile,"a") as otxmlfile:
    for x in range(len(df2)):
    # Check MAC address for other artefacts
        if re.fullmatch("[0-9a-fA-F]{12}$",df2["Mac-Addresses"].values[x]) is not None:
        #Put endpoints to XML ile 
            endpoint = f'<Endpoint macAddress="{df2["Mac-Addresses"].values[x]}" status="Known"><EndpointTags tagName="OT-Role" tagValue="{df2["Aruba User Role"].values[x]}"/>\
<EndpointTags tagName="Danfoss Approved Endpoint" tagValue="Yes"/></Endpoint>\n'
            otxmlfile.writelines(endpoint)
            #print (endpoint)
        else:
            # Check the MAC in file if MAC has artefacts
            print ("please recheck following MAC", df2["Mac-Addresses"].values[x],' in row',df2.index.values[x]+2)
    #write last(seconf) part of templateto OT XML file
    otxmlfile.writelines(template_lines[4:6])
    otxmlfile.close


