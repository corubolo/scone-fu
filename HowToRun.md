# Introduction #
Scone-fu contains tools to transform spreadsheets in Microsoft Excel format and the Open Document Format Spreadsheet (ODS) (used by Open Office/Libre Office and others) into XML which can be represented in a web browser through an XSL
stylesheet.

# Running the Excel to XML conversion, Excel feature extraction #

Given that the scone-fu source code has been downloaded, you can test the execute the conversion using Maven:
http://maven.apache.org/
After installing maven, you should be able to run the command that will compile and execute the Java applications for interaction with Excel spreadsheets:

## Excel to XML conversion ##

To execute the conversion on a full folder, you can call from the scone-fu source folder:

```
mvn package
mvn exec:java -Dexec.mainClass="uk.ac.liverpool.spreadsheet.example.TestToML" -Dexec.classpathScope=runtime -Dexec.args="./data/"
```

this will convert (leaving the original files intact) all the XLS,XLSX files in the "data" folder.
Each Excel file will be converted to multiple XML files, named as the original XLS file, postponed with "[number](sheet.md).xml"

## Excel feature extraction ##


To execute the featureextraction on a full folder, you can call from the scone-fu source folder:

```
mvn package
mvn exec:java -Dexec.mainClass="uk.ac.liverpool.spreadsheet.example.TestFA" -Dexec.classpathScope=runtime -Dexec.args="./data/"
```

this will analyse and report (leaving the original files intact) all the XLS,XLSX files in the "data" folder.
Each Excel file will generate one XML file named as the original Excel file, postponed with "_feature.xml"_

# Running the OpenOffice ODS conversion #
The Python script `convertODS2XML.py` is contained in the `code/python` directory. It makes use of the **odfpy** package that is also included and the **lxml** package (http://lxml.de/) that you should install in the standard Python location. Any Python version greater than 2.5 should be fine.

To run the script:
```
./convertODS2XML.py [-h][-v] <input.ods> <output.xml>
```
The input spreadsheet should be the first argument and the output xml file the last argument. At the moment the script will produce one xml file per sheet.