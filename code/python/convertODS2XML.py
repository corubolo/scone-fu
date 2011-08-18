#!/usr/bin/env python
# Script to read an openoffice spreadsheet and write out an html table
# including javascript for the formulae
# Adil Hasan, Univ of Liverpool
# License GPL

import re
import os
import getopt
import sys
import string

sys.path.append(sys.path[0] + '/odfpy-0.9.3')

try:
    import odf.opendocument
    import odf.table
    import odf.text
except ImportError:
    print 'Failed to import odf. Check PYTHONPATH'

try:
    import lxml.etree
except ImportError:
    print 'Failed to import lxml. Check module is in standard location'
    print 'or PYTHONPATH is set correctly.'

def usage():
    '''Function to print out the script usage
    '''
    print 'Script to read an openoffice spreadsheet input.ods and write out an html'
    print 'representation output.xhtml.'
    print 'Usage: convertODS2HTML.py [-h] <input.ods> <output.xhtml>'
    print ''

class Cell():
    '''Cell class
    '''
    def __init__(self):
        self.processed = 0
        self.id = ''
        self.colId = ''
        self.origFormula = ''
        self.formula = []
        self.value = '00'
        self.type = 'float'
        self.readOnly = '' 

def inLimit(ids, lower, upper, verboseFlag):
    '''Function to find whether a cell id is within lower and
    upper limits
    '''
    inside = False
    if (verboseFlag):
        print 'Lower limit: ', lower
        print 'Upper limit: ', upper
        print 'id: ', ids

    # If same row can use normal test
    idRow = int(re.findall('\d+', ids)[0])
    idCol = re.findall('[a-zA-Z]+', ids)[0]
    lowerRow = int(re.findall('\d+', lower)[0])
    lowerCol = re.findall('[a-zA-Z]+', lower)[0]
    upperRow = int(re.findall('\d+', upper)[0])
    upperCol = re.findall('[a-zA-Z]+', upper)[0]
    
    if (idRow == lowerRow):
        inside = (idCol >= lowerCol and idCol <= upperCol)
        return inside
    
    if (idRow == upperRow):
        inside = (idCol <= upperCol and idCol >= lowerCol)
        return inside

    # If not the same row we have to check each part separately
    if (idRow > lowerRow and idRow < upperRow 
        and idCol >= lowerCol and idCol <= upperCol):
        inside = True
    return inside

def processFormula(cells, formulaCellID, formula, verboseFlag):
    '''Function to process the formula and apply it to the cells.
    The formula is then removed from the target cell.
    '''
    colon = ':'
    
    startFormula = 'of:='
    newStart = '[%s]=' % (formulaCellID)
    
    # Parse the formula for the elements (we look for .X1 or
    # .X1;.X2 or .X1:.X2 all variables appear inbetween []
    elements = re.findall('\[(\.\w+\d+(\W\.\w+\d+)*)\]', formula)

    # New formula replaces the of:= with the acutal output cell
    newFormula = formula.replace(startFormula, newStart, 1)

    for element in elements:
        argument = element[0]
        if (verboseFlag):
            print 'argument is ', argument

        # If the arguments have ':' we have a range
        # Need to traverse range and apply the formula to all
        # cells within the range
        if (colon in argument):
            startVar, endVar = argument.split(colon)
            for rCells in cells:
                for aCell in rCells:
                    if (verboseFlag):
                        print 'newFormula ', newFormula
                        print 'inLimit ', inLimit(aCell.id, 
                                                  startVar, 
                                                  endVar,
                                                  verboseFlag)
                    if (inLimit(aCell.id, startVar, endVar, verboseFlag)):
                        aCell.formula.append(newFormula)
                        # Need to account for outputs that are inputs
                        # to another formula
                        if (aCell.origFormula and aCell.processed == 0):
                            aCell.processed = 0
                        else:
                            if (not aCell.origFormula):
                                aCell.readOnly = ''
                            aCell.processed = 1
        else:
            for rCells in cells:
                cellFound = False
                for aCell in rCells:
                    if (aCell.id == argument):
                        aCell.formula.append(newFormula)
                        # Need to account for outputs that are inputs
                        # to another formula
                        if (aCell.origFormula and aCell.processed == 0):
                            aCell.processed = 0
                        else:
                            if (not aCell.origFormula):
                                aCell.readOnly = ''
                            aCell.processed = 1
                        cellFound = True
                        break
                if (cellFound): break

    for rCells in cells:
        cellFound = False
        for aCell in rCells:
            if (aCell.id == formulaCellID):
                aCell.processed = 1
                cellFound = True
                break
        if (cellFound): break
    return cells

class SpreadSheet():
    '''Spreadsheet class
    '''
    def __init__(self):
        self.colHeading = []
        self.root = lxml.etree.Element("spreadsheets")
        self.cell = None
        self.cols = []

    def insertTable(self, table):
        '''Method to insert the table into the spreadsheet
        '''
        self.table = lxml.etree.SubElement(self.root, "Table",
                                          name=str(table.getAttribute("name")))

    def storeCol(self, col):
        '''Store the column name
        '''
        if (col not in self.cols):
            self.cols.append(col)
    
    def insertRowHeader(self, cells, aCell):
        '''Insert data into a row header
        '''
        self.cell = lxml.etree.SubElement(cells, "RowHeader",
                                          value_type=aCell.type,
                                          value=aCell.value,
                                          readOnly=aCell.readOnly)
        self.cell.text = str(aCell.value)

    def insertCell(self, cells, aCell):
        '''Insert data into a table cell
        '''
        forin = ''
        kwArgs = {}

        if (len(aCell.formula) > 1):
            firstPass = True
            for f in aCell.formula:
                if (not firstPass):
                    forin += ' || %s' % f
                else:
                    forin = f
                    firstPass = False
        elif (len(aCell.formula) == 1):
            forin = aCell.formula[0]
        
        if (len(aCell.type) > 0):
            kwArgs['value_type'] = aCell.type
        if (len(aCell.value) > 0):
            kwArgs['value'] = aCell.value
        if (len(aCell.formula) > 0):
            kwArgs['formula'] = forin
        if (len(aCell.id) > 0):
            kwArgs['cellID'] = aCell.id
        if (len(aCell.origFormula) > 0):
            origFormula = aCell.origFormula.split("of:=")[1] 
            kwArgs['cellFormula'] = origFormula
        if (len(aCell.readOnly) > 0):
            kwArgs['readOnly'] = aCell.readOnly

        self.cell = lxml.etree.SubElement(cells, "TableCell", kwArgs)
        
        if (self.cell is not None):
            self.cell.text = str(aCell.value)
    
    def insertHeader(self, headers, heading):
        '''Insert the header
        '''
        self.header = lxml.etree.SubElement(headers, "ColumnHeader")
        self.header.text = heading
        self.colHeading.append(heading)

    def insertHeaders(self):
        '''Insert the headers list
        '''
        self.headers = lxml.etree.SubElement(self.table, "ColumnHeaders")

    def insertCells(self, row):
        '''Insert the cells list
        '''
        self.cells = lxml.etree.SubElement(row, "TableCells")

    def insertRow(self):
        '''Insert a row into the table
        '''
        self.row = lxml.etree.SubElement(self.table, "TableRow")
        
    def serialize(self):
        '''serialize the xml as a string
        '''
        return lxml.etree.tostring(self.root, encoding='iso-8859-1',
                                   xml_declaration=True, 
                                   pretty_print=True)

def getCellID(col, row):
    '''Function to return the cell ID
    '''
    count = col/26
    indx = col%26
    cellPrefix = ''
    cellPrefix = string.ascii_uppercase[0]*(count)
    colChar = cellPrefix + string.ascii_uppercase[indx]
    cellID = ".%s%s" % (colChar, row)
    return cellID, colChar

def fillCell(aCell, cell, col, row):
    '''Function to fill a spreadsheet cell attributes
    '''
    value_type = ''
    value = ''
    formula = ''
    cellID, colID = getCellID(col, row)
    aCell.id = cellID
    aCell.colId = colID
    
    for attrib in cell.attributes.keys():
        if ('value-type' in attrib):
            aCell.type = cell.attributes[attrib]
        if ('value' in attrib):
            aCell.value = cell.attributes[attrib]
        if ('formula' in attrib):
            aCell.origFormula = cell.attributes[attrib]
            aCell.readOnly = 'readOnly'
    # In the case of a cell with a string we have child nodes that
    # contain the value
    if (aCell.type == "string" and aCell.value == '00'):
        aCell.value = str(cell.childNodes[0].childNodes[0])

    return aCell, colID
 
def processSheet(docObj, spreadSheet, verboseFlag):
    '''Method to process the rows from the spreadsheet
    '''
    
    rows = docObj.getElementsByType(odf.table.TableRow)
    rowCount = 1
    repeatedCols = 'number-columns-repeated'
    sCells = []
    for row in rows:
        cells = row.getElementsByType(odf.table.TableCell)
        colCount = 0
        rCells = []
        for cell in cells:
            repeated = 0
            if (verboseFlag):
                print 'attribs ', cell.attributes
            for attrib in cell.attributes.keys():
                # If we have a repeated cell we need to apply the
                # correct values and then skip to the next unique cell
                if (repeatedCols in attrib):
                    for colc in range(0, int(cell.attributes[attrib])):
                        aCell = Cell()
                        aCell, colID = fillCell(aCell, cell, 
                                                colCount, rowCount)
                        if (verboseFlag):
                            print 'id ', colID, ' val ', aCell.value
                        rCells.append(aCell)
                        spreadSheet.storeCol(colID)
                        repeated = 1
                        colCount += 1
                    break
            
            if (not repeated):
                aCell = Cell()
                aCell, colID = fillCell(aCell, cell, 
                                        colCount, rowCount)
                if (verboseFlag):
                    print 'id ', colID, ' val ', aCell.value
                rCells.append(aCell)
                spreadSheet.storeCol(colID)
                colCount += 1

        rowCount+= 1
        sCells.append(rCells)

    return spreadSheet, sCells

def fillMissingCells(sCells, spreadSheet, verboseFlag):
    '''Method to find missing cells and fill them in
    '''    
    tsCells = []
    rowC = 1
    for rCells in sCells:
        trCells = []
        for col in spreadSheet.cols:
            cellIn, cell = cellInRow(col, rCells)
            if (cellIn):
                if (verboseFlag):
                    print 'cellIn: cell.id %s cell.origFormula %s \n' \
                            % (cell.id, cell.origFormula)
                trCells.append(cell)
                continue
            else:
                if (verboseFlag):
                    print "cellOut: cell.id .%s%s" % (col, rowC) 
                aTCell = Cell()
                aTCell.id = ".%s%s" %(col, rowC)
                trCells.append(aTCell)
        tsCells.append(trCells)
        rowC += 1
    
    sCells = tsCells

    return sCells

def applyFormulaToCells(sCells, verboseFlag):
    '''Function to apply the formula to the relevant cells
    '''
    nCells = sCells
    while (1):
        processedFormula = False
        for rCells in sCells:
            for aCell in rCells:
                if (aCell.origFormula and aCell.processed == 0):
                    if (verboseFlag):
                        print 'aCell.id %s aCell.origFormula %s' \
                                % (aCell.id, aCell.origFormula)
                    nCells = processFormula(nCells, 
                                            aCell.id, 
                                            aCell.origFormula, 
                                            verboseFlag)
                    processedFormula = True
                    break

        if (not processedFormula):
            break
        sCells = nCells

    return sCells

def convertDocument(stylesheet, inputFile, outputFile, verboseFlag):
    '''Function to convert the document
    '''
    # Load the document into memory
    doc = odf.opendocument.load(inputFile)
    
    # Treat it as a spreadsheet
    docObj = doc.spreadsheet

    # Create an instance of the spreadsheet XML file
    spreadSheet = SpreadSheet()
    
    # Get the sheets
    sheets = docObj.getElementsByType(odf.table.Table)

    # Loop over the sheets and process the cells
    sheetCount = 0
    for aSheet in sheets:
        spreadSheet.insertTable(aSheet)

        # Get the table rows from the document and process each row
        spreadSheet, sCells = processSheet(aSheet, spreadSheet, verboseFlag)
    
        # Loop over the number of cols and construct the headers
        spreadSheet.insertHeaders()
        spreadSheet.insertHeader(spreadSheet.headers, 'RowID\ColID')
        for col in spreadSheet.cols:
            spreadSheet.insertHeader(spreadSheet.headers, col)
     
        # Check for missing cells and fill them in
        sCells = fillMissingCells(sCells, spreadSheet, verboseFlag)
  
        # Loop over the cells and apply the formula to the cells
        sCells = applyFormulaToCells(sCells, verboseFlag)
      
        # Insert the cells into the xml file
        rowCount = 1
        for rCells in sCells:
            hCell = Cell();
            hCell.type = "rowID"
            hCell.value = str(rowCount)
            spreadSheet.insertRow()
            spreadSheet.insertRowHeader(spreadSheet.row, hCell)
            spreadSheet.insertCells(spreadSheet.row)

            for aCell in rCells:
                spreadSheet.insertCell(spreadSheet.cells, aCell)
            rowCount += 1
 
        # For now we output the spreadsheets into different output files
        # until we can handle multiple sheets in the same file
        foutNames = os.path.split(outputFile)
        baseNames  = foutNames[1].split(os.path.extsep)
        newOutFile = "%s-%s" % (baseNames[0], sheetCount)
        
        for fout in baseNames[1:]:
            newOutFile += "%s%s" % (os.path.extsep, fout)
        
        newOutFile = os.path.join(foutNames[0], newOutFile)
        
        fo = file(newOutFile, 'w')
        xmlOut = spreadSheet.serialize().split('\n')
        fo.write("%s\n" % xmlOut[0])
        fo.write("%s\n" % stylesheet)
        for i in range(1, len(xmlOut)):
            fo.write("%s\n" % xmlOut[i])
        fo.close()

        # Start a new spreadsheet (for now - in the future should all be
        # in one file)
        sheetCount += 1
        spreadSheet = SpreadSheet()

def cellInRow(col, cells):
    '''Function to find if a cell is contained in a row
    '''
    inside = False
    cell = None
    for cell in cells:
        if (col == cell.colId):
            inside = True
            break
    return inside, cell

if __name__ == '__main__':

    opts, args = getopt.getopt(sys.argv[1:], 'hv', ['help', 'verbose'])
    if (len(args) != 2):
        print 'Error: Input and output file must be specified!\n\n'
        usage()
        sys.exit()
    else:
        inputFile = args[0].strip()
        if (not(os.path.isfile(inputFile))):
            print 'Error: Cannot open file %s' % inputFile
            sys.exit(2)
        outputFile = args[1].strip()

    verboseFlag = False
    stylesheet = '<?xml-stylesheet type="text/xsl" href="spreadsheet.xsl"?>'

    for opt, val in opts:
        if (opt == '-h' or opt == '--help'):
            usage()
            sys.exit(0)
        if (opt == '-v' or opt == '--verbose'):
            verboseFlag = True
    
    convertDocument(stylesheet, inputFile, outputFile, verboseFlag)
