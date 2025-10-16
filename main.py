import src.xls2xml as bibXml

# Initialize the class with parameters
# * inputFile = relative path to the XLS file
# * marcxml => Boolean (True => Output XML has MARC namespace)
xmlFile = bibXml.CreateBibXML(inputFile="boilerplate/AlmaUpdate_ARKUrl_PLSa81ff_20251015.xlsx",marcxml=False)

# Loop through the Excel input file, specifically the pandas DataFrame
for index, row in xmlFile.df.iterrows():
    # Call the method to generate the XML record element
    xmlFile.addRecord(row)
    #break

# Write XML-File
xmlFile.writeXmlFile()

