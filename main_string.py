import pandas as pd
import os

# --- CONFIGURATION ---
excel_file = "HSS drill mass upload.xlsx"   # Your Excel file
output_folder = "xml_output"  # Folder to store XML files
id_column = "Name"  # Column to use for file naming (can be None to use row index)

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Read Excel file but skip the first 7 rows
df = pd.read_excel(excel_file, skiprows=11)

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    
    # Determine file name
    if id_column and id_column in df.columns:
        clean = row[id_column].replace("/","-")
        clean = clean.replace('"',"in")
        file_name = f"{row[id_column]}.xml"
    else:
        file_name = f"row_{index+1}.xml"

    #row.Name.replace("\"","&quot;")
    #quoteEscapedName = row.Name
    #df['Name'] = df['Name'].str.replace('\"', '&quot;')
    row.Name = row.Name.replace('”', '&quot;')
    row.Name = row.Name.replace('"', '&quot;')


    xml_string = f"""<?xml version="1.0" encoding="UTF-8"?>
<omtdx version="34" srcURL="sqlite://O:\Mech &amp; Fluid Systems\Equipment\Fabrication\Mill_CNC_Doosan6500\HyperMILL Files\Tool Database\DVF6500.db" srcGuid="4a03a81b-a90d-4a23-9f7a-383bffba2be4" srcType="standard" files="HSS Twist _files">
<tools>
    <tools folder="Drills" objGuid="07a1b934-78d2-4c38-b2ba-6fafb61fef52">
    <tools folder="Drills (Solid)" objGuid="e4b360d5-bddc-471a-bb6c-993839ca7b68">
        <tools folder="HSS" objGuid="dbf99580-3c66-4f81-9ac9-e30844c191fa">
        <tool type="drilTool" name="{row.Name}">
            <param name="hmbState" value="unlinked"/>
            <param name="comment" value="{row.Comment2}"/>
            <param name="objGuid" value="2184a13d-41a4-4ccb-b860-38b8fbe07d3e"/>
            <param name="cuttingMaterial" value="HSS"/>
            <param name="lengthOfUnit" value="inch"/>
            <param name="toolTotalLength" value="{row.totalToolLength}"/>
            <param name="spindleRotation" value="clockwise"/>
            <coupling type="shank">
            <param name="hmbState" value="unlinked"/>
            <param name="objGuid" value="0e33ab58-a401-45d2-81d4-1ff60c74eac8"/>
            <param name="minDia" value="{row.minDia_toolDiameter}"/>
            <param name="minLen" value="0"/>
            <param name="maxLen" value="0"/>
            <param name="lengthOfUnit" value="inch"/>
            </coupling>
            <param name="cuttingEdges" value="1"/>
            <param name="cuttingLength" value="{row.cuttingLength}"/>
            <param name="toolShaftType" value="none"/>
            <param name="toolDiameter" value="{row.minDia_toolDiameter}"/>
            <tecsets>
            <tecset type="milling">
                <param name="lengthOfUnit" value="inch"/>
                <param name="planeFeedrate" value="8"/>
                <param name="coolants" value="1"/>
                <param name="spindleSpeed" value="2000"/>
                <param name="cuttingSpeed" value="10"/>
                <param name="zFeedrate" value="2"/>
                <param name="feedratePerEdge" value="0.02"/>
                <param name="drillingFeedrate" value="{row.drillingFeedrate}"/>
                <param name="cuttingWidth" value="0"/>
                <param name="cuttingLength" value="0"/>
                <param name="maxRedFeedrateAngle" value="15"/>
                <param name="plungeAngle" value="2"/>
                <param name="reducedFeedrate" value="4"/>
            </tecset>
            <tecset type="milling">
                <param name="material" value="Wrought aluminium alloys Si &lt; 6%"/>
                <param name="purpose" value="Drilling"/>
                <param name="lengthOfUnit" value="inch"/>
                <param name="planeFeedrate" value="0"/>
                <param name="coolants" value="1"/>
                <param name="spindleSpeed" value="1"/>
                <param name="cuttingSpeed" value="200"/>
                <param name="zFeedrate" value="1"/>
                <param name="feedratePerEdge" value="0"/>
                <param name="drillingFeedrate" value="{row.drillingFeedrate}"/>
                <param name="cuttingWidth" value="0"/>
                <param name="cuttingLength" value="0"/>
                <param name="maxRedFeedrateAngle" value="0"/>
                <param name="plungeAngle" value="0"/>
                <param name="reducedFeedrate" value="0"/>
                <param name="spindleSpeedFormula" value="fSinch"/>
                <param name="zFeedrateFormula" value="fFHSSDrill"/>
            </tecset>
            <tecset type="milling">
                <param name="material" value="Titanium alloys &gt; 300 HB [Ti6Al4V]"/>
                <param name="purpose" value="Drilling"/>
                <param name="lengthOfUnit" value="inch"/>
                <param name="planeFeedrate" value="0"/>
                <param name="coolants" value="1"/>
                <param name="spindleSpeed" value="1"/>
                <param name="cuttingSpeed" value="20"/>
                <param name="zFeedrate" value="1"/>
                <param name="feedratePerEdge" value="0"/>
                <param name="drillingFeedrate" value="{row.drillingFeedrate}"/>
                <param name="cuttingWidth" value="0"/>
                <param name="cuttingLength" value="0"/>
                <param name="maxRedFeedrateAngle" value="0"/>
                <param name="plungeAngle" value="0"/>
                <param name="reducedFeedrate" value="0"/>
                <param name="spindleSpeedFormula" value="fSinch"/>
                <param name="zFeedrateFormula" value="fFHSSDrill"/>
            </tecset>
            <tecset type="milling">
                <param name="material" value="Steel 500 - 850 N/mm²  (Steel up to 24 HRC)"/>
                <param name="purpose" value="Drilling"/>
                <param name="lengthOfUnit" value="inch"/>
                <param name="planeFeedrate" value="0"/>
                <param name="coolants" value="1"/>
                <param name="spindleSpeed" value="1"/>
                <param name="cuttingSpeed" value="50"/>
                <param name="zFeedrate" value="0"/>
                <param name="feedratePerEdge" value="0"/>
                <param name="drillingFeedrate" value="{row.drillingFeedrate}"/>
                <param name="cuttingWidth" value="0"/>
                <param name="cuttingLength" value="0"/>
                <param name="maxRedFeedrateAngle" value="0"/>
                <param name="plungeAngle" value="0"/>
                <param name="reducedFeedrate" value="0"/>
                <param name="spindleSpeedFormula" value="fSinch"/>
                <param name="zFeedrateFormula" value="fFHSSDrill"/>
            </tecset>
            <tecset type="milling">
                <param name="material" value="Stainless steel ferritic/martensitic"/>
                <param name="purpose" value="Drilling"/>
                <param name="lengthOfUnit" value="inch"/>
                <param name="planeFeedrate" value="0"/>
                <param name="coolants" value="1"/>
                <param name="spindleSpeed" value="1"/>
                <param name="cuttingSpeed" value="30"/>
                <param name="zFeedrate" value="0"/>
                <param name="feedratePerEdge" value="0"/>
                <param name="drillingFeedrate" value="{row.drillingFeedrate}"/>
                <param name="cuttingWidth" value="0"/>
                <param name="cuttingLength" value="0"/>
                <param name="maxRedFeedrateAngle" value="0"/>
                <param name="plungeAngle" value="0"/>
                <param name="reducedFeedrate" value="0"/>
                <param name="spindleSpeedFormula" value="fSinch"/>
                <param name="zFeedrateFormula" value="fFHSSDrill"/>
            </tecset>
            </tecsets>
            <param name="coneAngle" value="118"/>
            <param name="centeringRequired" value="0"/>
            <param name="breakThroughLength" value="{row.breakThroughLength}"/>
        </tool>
        </tools>
    </tools>
    </tools>
</tools>
<ncTools>
    <ncTools folder="Solid Drills" objGuid="361e01b2-a8a7-41ed-a682-09b3f947be51">
    <ncTools folder="HSS Twist" objGuid="d5604499-975e-49bf-994a-34656da74d32">
        <ncTool ncNumber="31" id="import test" name="HSS Twist">
        <param name="hmbState" value="unlinked"/>
        <param name="comment" value="{row.Comment}"/>
        <param name="objGuid" value="65d25020-e7ad-4a31-bb43-e099ef18926a"/>
        <param name="lengthOfUnit" value="inch"/>
        <param name="toolLength" value="0"/>
        <param name="usableLength" value="0"/>
        <param name="clearanceLength" value="0"/>
        <param name="gageLength" value="0"/>
        <param name="spindleSpeedFactor" value="1"/>
        <param name="feedrateFactor" value="1"/>
        <param name="cuttingWidthFactor" value="1"/>
        <param name="cuttingLengthFactor" value="1"/>
        <param name="maxSpindleSpeed" value="0"/>
        <param name="maxFeedrate" value="0"/>
        <param name="breakageCheck" value="0"/>
        <components>
            <component type="spindle" name="DVF6500" contour="Spindle_1"/>
            <component type="holder" name="NCAT40-NPU13-105U-IDU" contour="NCAT40-NPU13-105U-IDU" reach="4.7795276641845703"/>
            <component type="tool" name="{row.Name}" reach="{row.reach}"/>
        </components>
        <cuttingProfiles>
            <cuttingProfile type="milling">
            <param name="enabled" value="1"/>
            </cuttingProfile>
            <cuttingProfile material="Wrought aluminium alloys Si &lt; 6%" purpose="Drilling" type="milling">
            <param name="enabled" value="1"/>
            </cuttingProfile>
            <cuttingProfile material="Steel 500 - 850 N/mm²  (Steel up to 24 HRC)" purpose="Drilling" type="milling">
            <param name="enabled" value="1"/>
            </cuttingProfile>
            <cuttingProfile material="Stainless steel ferritic/martensitic" purpose="Drilling" type="milling">
            <param name="enabled" value="1"/>
            </cuttingProfile>
            <cuttingProfile material="Titanium alloys &gt; 300 HB [Ti6Al4V]" purpose="Drilling" type="milling">
            <param name="enabled" value="1"/>
            </cuttingProfile>
        </cuttingProfiles>
        <coupling type="isoAdaptor">
            <param name="hmbState" value="unlinked"/>
            <param name="objGuid" value="3619846f-bc46-4c91-b46e-983d7cfc4e3f"/>
            <param name="class" value="CAT 40"/>
            <param name="isoCode" value="SKG3*C0400$$$$"/>
            <param name="lengthOfUnit" value="inch"/>
        </coupling>
        </ncTool>
    </ncTools>
    </ncTools>
</ncTools>

</omtdx>
    """
    
    file_name = file_name.replace('"',"in")
    file_name = file_name.replace('”','in')
    file_name = file_name.replace("/","_")
    
     # Write XML to file for this record
    output_path = os.path.join(output_folder, file_name)
    # Writing a string to a file
    with open(output_path, "w") as file:
        file.write(xml_string)
    

print(f"XML files created in: {output_folder}")