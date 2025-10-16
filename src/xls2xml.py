import pandas as pd
import lxml.etree as ET
import xml.dom.minidom as minidom
import re


class CreateBibXML:
    # Initialisiere Klasse CreateBibXML
    # Ein xls-InputFile wird übergeben
    # Wenn Output als MARC-XML übergeben wird, wird hier der Namespace definiert und das Root-Element erzeugt
    # MARC==False erzeugt ein XML im Stil des erwarteten Formats für Alma-Import-Profile-XML (marc:record wird durch <collection> ersetzt)
    def __init__(self, inputFile, marcxml=True):
        if marcxml == True:
            # Define namespace
            self.namespace = "http://www.loc.gov/MARC21/slim"
            ET.register_namespace("marc", self.namespace)
            self.root = ET.Element(f"{{{self.namespace}}}records")
        else:
            self.root = ET.Element(f"collection")
            self.namespace = ""
        
        self.inputFileName = f"{inputFile}"
        self.outputFileName = re.sub("xlsx$","xml",self.inputFileName)
        self.data = pd.read_excel(self.inputFileName, dtype=str)
        self.df = pd.DataFrame(self.data)
        #print (self.df.head())
        if 'leader' not in self.df.columns:
            self.df.insert(0,'leader', '00000nam#a2200000#u#4500' )
        self.initColumns = self.df.columns

    #Helper Method to extract fieldnumber, indices and subfield codes from excel-column-header
    def getTagName(self, tagString):
        regex = r"tag([\d]{3})|([\d\s]{1})([\d\s]{1})\$([a-z0-9]{1})"
        matches = re.finditer(regex, tagString, re.MULTILINE)
        respond={"tag":"","ind1":"","ind2":"","sub":""}
        for matchNum, match in enumerate(matches, start=1):
            if match.group(2) != None:
                respond["ind1"] = match.group(2).strip()
                respond["ind2"] = match.group(3).strip()
                respond["sub"] = match.group(4)
            else:
                respond["tag"] = match.group(1)
        return respond


    def addRecord(self, row):
        #Create new (marc:)record node
        record = ET.SubElement(self.root, f"{{{self.namespace}}}record")
        current_columns = row.index.tolist()

        for col in current_columns:
            if pd.notna(getattr(row, col)):  # Only add the value if it's not NA
                if col.startswith('tag'):
                    tagNames = self.getTagName(col)
                    # print(tagNames)
                    # print(len(tagNames))
                    print(f"add marc field {col}: {row[col]} ({tagNames['tag']})")
                    if (int(tagNames["tag"])<10):
                        fieldType = "controlfield"
                    else:
                        fieldType = "datafield"
                    
                    #Check if Field by TagName already set in ET
                    datafield = record.find(f".//*[@tag='{tagNames['tag']}']")
                    if datafield == None:
                        if fieldType == "datafield":
                            datafield = ET.SubElement(record, f"{{{self.namespace}}}{fieldType}", tag = tagNames["tag"], ind1=tagNames["ind1"], ind2=tagNames["ind2"])
                        else:
                            datafield = ET.SubElement(record, f"{{{self.namespace}}}{fieldType}", tag = tagNames["tag"])

                    if tagNames["sub"] == "":
                        datafield.text = row[col]
                    else:
                        subfield = ET.SubElement(datafield, f"{{{self.namespace}}}subfield", code=tagNames["sub"])
                        subfield.text = row[col]

                elif col.startswith("leader"):
                    # print(f"add {col}: {row[col]}")
                    datafield = ET.SubElement(record, f"leader")
                    datafield.text = row[col]
        
        # Separate leader from other children
        leader = record.find("leader")
        others = [el for el in record if el.tag != "leader"]
                
        # Sort child nodes by numeric 'tag' attribute
        sorted_others = sorted(
            others,
            key=lambda el: int(el.get("tag")) if el.get("tag") and el.get("tag").isdigit() else float("inf")
        )

        # Clear and re-append in correct order
        record[:] = []  # remove all children
        if leader is not None:
            record.append(leader)
        for el in sorted_others:
            record.append(el)



    def writeXmlFile(self, outputfile=""):
        if outputfile == "":
            outputfile = self.outputFileName
        # Generate XML tree
        tree = ET.ElementTree(self.root)

        # Convert to pretty XML format
        xml_str = ET.tostring(self.root, encoding="unicode")
        pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")

        # Save to a file
        with open(outputfile, "w", encoding="utf-8") as xml_file:
            xml_file.write(pretty_xml)

        print(f"XML file created as {outputfile}")    

