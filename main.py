import pandas as pd
import os
import xml.etree.ElementTree as ET

# --- CONFIGURATION ---
excel_file = "input.xlsx"   # Your Excel file
output_folder = "xml_output"  # Folder to store XML files
id_column = "Name"  # Column to use for file naming (can be None to use row index)

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Read Excel file
df = pd.read_excel(excel_file)

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    # Determine file name
    if id_column and id_column in df.columns:
        file_name = f"{row[id_column]}.xml"
    else:
        file_name = f"row_{index+1}.xml"
    
    # Create XML root element
    root = ET.Element("Record")
    
    # Add each column as a child element
    for col_name, value in row.items():
        
        print(f"{col_name},{value}")
        #print(f"{col_name}")
        #print(f"{value}")
        child = ET.SubElement(root, col_name)
        child.text = "" if pd.isna(value) else str(value)
    
    # Build XML tree
    tree = ET.ElementTree(root)
    
    # Write XML to file
    output_path = os.path.join(output_folder, file_name)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

print(f"XML files created in: {output_folder}")