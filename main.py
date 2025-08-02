import pandas as pd
import os
import xml.etree.ElementTree as ET

# --- CONFIGURATION ---
excel_file = "HSS drill mass upload.xlsx"   # Your Excel file
output_folder = "xml_output"  # Folder to store XML files
id_column = "Name"  # Column to use for file naming (can be None to use row index)

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Read Excel file but skip the first 7 rows
df = pd.read_excel(excel_file, skiprows=7)

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    
    # Determine file name
    if id_column and id_column in df.columns:
        file_name = f"{row[id_column]}.xml"
    else:
        file_name = f"row_{index+1}.xml"
    
    # Create XML root element
    attribs={"version": "23", 
             "srcType": "standard", 
             "match-formulas-by-expression" :"true",
             "match-material-by-provider" : "fraisa",
             "tecset-values-outdated":"true"}
    root = ET.Element('omtdx', attrib=attribs)

    #technologyPurposes
    technologyPurposes = ET.SubElement(root, "technologyPurposes")
    technologyPurpose = ET.SubElement(technologyPurposes, "technologyPurpose")
    
    #Formulas
    formulas = ET.SubElement(root, "formulas")
    formula = ET.SubElement(formulas, "formula",attrib={"name":"fFDrill", "type":"feedrate"})
    param = ET.SubElement(formula, "param", attrib={"name":"formula", "value":"f*n"})
    formula = ET.SubElement(formulas, "formula",attrib={"fS":"fFDrill", "type":"speed"})
    param = ET.SubElement(formula, "param", attrib={"name":"formula", "value":"(Vc*1000)/(d*pi)"})
    
    #Matrials
    materials = ET.SubElement(root, "materials")
    material = ET.SubElement(materials, "material",attrib={"name":f"{row.Name}"})

    #cuttingMaterials
    cuttingMaterials = ET.SubElement(root, "cuttingMaterials")
    cuttingMaterial = ET.SubElement(cuttingMaterials, "cuttingMaterial",attrib={"name":f"{row.Name}"})

    #couplings
    couplings = ET.SubElement(root, "couplings")

    #coolants
    coolants = ET.SubElement(root, "coolants")
    coolant = ET.SubElement(coolants, "coolant",attrib={"number":"1"})
    param = ET.SubElement(coolant, "param", attrib={"name":"comment", "value":"External coolant"})
    param = ET.SubElement(coolant, "param", attrib={"name":"type", "value":"external"})
    
    #tools
        #tool
            #tecsets
                #tecset
    tools = ET.SubElement(root, "tools")
    tool = ET.SubElement(tools, "tool",attrib={"type":"drilTool", "name":"1"})
    param = ET.SubElement(tool, "param", attrib={"name":"comment", "value":f"{row.Comment}"})
    param = ET.SubElement(tool, "param", attrib={"name":"orderingCode", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"manufacturer", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"cuttingMaterial", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"lengthOfUnit", "value":f"{row.overall_length_in}"})
    param = ET.SubElement(tool, "param", attrib={"name":"toolTotalLength", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"cuttingEdges", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"cuttingLength", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"toolShaftType", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"toolShaftDiameter", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"toolShaftChamferDefMode", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"toolShaftChamferAbsPos", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"toolDiameter", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"taperHeight", "value":f"unknown"})
    param = ET.SubElement(tool, "param", attrib={"name":"coneAngle", "value":f"unknown"})
    tecsets = ET.SubElement(tool, "tecsets")
    tecset = ET.SubElement(tecsets,"tecset", attrib={"type":"milling"})
    param = ET.SubElement(tecset,"param", attrib={"cuttingMaterial":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"material":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"purpose":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"lengthOfUnit":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"spindleSpeedFormula":"fS"})
    param = ET.SubElement(tecset,"param", attrib={"cuttingSpeed":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"coolants":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"zFeedrateFormula":"fFDrill"})
    param = ET.SubElement(tecset,"param", attrib={"reducedFeedrateFormula":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"plungeAngle":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"maxRedFeedrateAngle":f"unknown"})
    param = ET.SubElement(tecset,"param", attrib={"drillingFeedrate":f"unknown"})
   
    # Build XML tree and format it
    tree = ET.ElementTree(root)
    ET.indent(tree, space="\t", level=0)
    
    file_name = file_name.replace('"',"in")
    file_name = file_name.replace('”','in')
    file_name = file_name.replace("/","_")
    
     # Write XML to file for this record
    output_path = os.path.join(output_folder, file_name)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

print(f"XML files created in: {output_folder}")