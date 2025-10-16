# XLS2XML Writer for bibliographic records

Lightweight python command-line tool to create xml (generic xml or marc:xml) files based on xls-spreadsheets. Those xml files could be used e.g. the run a file-based import profile in alma to create or update bibliographic records.

## Requirements
The script needs pandas and openpyxl for dealing with the xls-file, and lxml for the xml manipulation. You can run on your environment:

```bash
pip install -r requirements.txt
```

## Basic Usage

### Definition of input spreadsheet

To run the file you can prepare a spreadsheet the following way. Column headers has the syntax
`tag{FieldNumber}{Ind1}{Ind2}${subfieldcode}`
* Fieldnumber is always three characters long
* Index1 and Index2 are given as number (0-9) or even as Spatium \s (e.g. `tag0247 $a` will be transformed in marc field 024 with ind1=7 and ind2 is blank)
* Subfield as one character (a-z or 0-9) always with prefix $ in column header.

| **tag001**         | **tag85642$u**                                | **tag85642$3**                     |
|--------------------|-----------------------------------------------|------------------------------------|
| 9911............01 | https://test.url                              | My Test URL Description in $$3     |

### Run the script

Open `main.py` and add the path to your input xls-File and decide if you need marc-xml or non marc-xml output. (default is marc-xml)
If you don't need any changes in the data itself, the `main.py` already has implemented a loop throught the dataframe, with the method of transforming the tabular data into a bibliographic xml style and write the xml-file to the same path as the input-file was given. Method `writeXmlFile(outputfile="")` has an optional paramter for the path of the outputfile if desired. 

```python
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
```

### Add code to transform your bibliographic data

Within the loop `for index, row in xmlFile.df.iterrows():` you can add code as necessary to transform your data. 

```python
for index, row in xmlFile.df.iterrows():
    # Here add your code for individual  processing on the data fields.
    # For example, create a new field based on an existing value, or standardise content.
    # The given example removes a part of an url to create the content for an identifier-marc-field (024)
    import re
    ark = re.sub("https://n2t.net/","",row["tag85642$u"])
    row["tag0247 $a"] = ark
    row["tag0247 $2"] = "ark"
    
    # Aufruf der Methode zur Erzeugung des xml-Record-Elements
    xmlFile.addRecord(row)

    #break
```

### Output

The example data table shown above leads with paramter `marc=False` to the following xml-result:

```xml
<?xml version="1.0" ?>
<collection>
  <record>
    <leader>00000nam#a2200000#u#4500</leader>
    <controlfield tag="001">9911............01</controlfield>
    <datafield tag="856" ind1="4" ind2="2">
      <subfield code="u">https://test.url</subfield>
      <subfield code="3">My Test URL Description in $$3 </subfield>
    </datafield>
  </record>
</colleciton>
```
Output with paramter `marc=True`

```xml
<?xml version="1.0" ?>
<marc:records xmlns:marc="http://www.loc.gov/MARC21/slim">
  <marc:record>
    <leader>00000nam#a2200000#u#4500</leader>
    <marc:controlfield tag="001">9911............01</marc:controlfield>
    <marc:datafield tag="856" ind1="4" ind2="2">
      <marc:subfield code="u">https://test.url</marc:subfield>
      <marc:subfield code="3">My Test URL Description in $$3</marc:subfield>
    </marc:datafield>
  </marc:record>
```