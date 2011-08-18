/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

// modified to fix updates to the POI library
package org.apache.poi.ss.examples.html;

import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.Formatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.AreaPtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import uk.ac.liverpool.spreadsheet.FormulaRenderer;

/**
 * This example shows how to export a spreadsheet in XML using the classes for
 * spreadsheet display. This includes formulae transformation .
 * 
 * Modified by Fabio Corubolo
 * 
 * @author Ken Arnold, Industrious Media LLC
 * @author Fabio Corubolo, University of Trento
 */
public class ToXML {
    private final Workbook wb;
    private final Appendable output;
    private Formatter out;
    private boolean gotBounds;
    private int firstColumn;
    private int endColumn;
    private int currentSheet;
    private HSSFWorkbook hswb;
    private XSSFWorkbook xswb;
    
    private boolean evaluateFormulae= false;

    private Map<Integer, String> colNumbers;

    private Map<String, List<String>> cellsToFormula;

    private Map<String, String> crToParent;

    private Map<String, String> cellToFormulaConverted;

    /**
     * Creates a new converter to XML for the given workbook.
     * 
     * @param wb
     *            The workbook.
     * @param output
     *            Where the HTML output will be written.
     * 
     * @return An object for converting the workbook to HTML.
     */
    public static ToXML create(Workbook wb, Appendable output) {
        return new ToXML(wb, output);
    }

    /**
     * Creates a new converter to HTML for the given workbook. If the path ends
     * with "<tt>.xlsx</tt>" an {@link XSSFWorkbook} will be used; otherwise
     * this will use an {@link HSSFWorkbook}.
     * 
     * @param path
     *            The file that has the workbook.
     * @param output
     *            Where the HTML output will be written.
     * 
     * @return An object for converting the workbook to HTML.
     */
    public static ToXML create(String path, Appendable output)
            throws IOException {
        return create(new FileInputStream(path), output);
    }

    /**
     * Creates a new converter to HTML for the given workbook. This attempts to
     * detect whether the input is XML (so it should create an
     * {@link XSSFWorkbook} or not (so it should create an {@link HSSFWorkbook}
     * ).
     * 
     * @param in
     *            The input stream that has the workbook.
     * @param output
     *            Where the HTML output will be written.
     * 
     * @return An object for converting the workbook to HTML.
     */
    public static ToXML create(InputStream in, Appendable output)
            throws IOException {
        try {
            Workbook wb = WorkbookFactory.create(in);
            return create(wb, output);
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException(
                    "Cannot create workbook from stream", e);
        }
    }

    private ToXML(Workbook wb, Appendable output) {
        if (wb == null)
            throw new NullPointerException("wb");
        if (output == null)
            throw new NullPointerException("output");
        this.wb = wb;
        this.output = output;

        if (wb instanceof HSSFWorkbook) {
            hswb = (HSSFWorkbook) wb;
        } else if (wb instanceof XSSFWorkbook) {
            xswb = (XSSFWorkbook) wb;
        }
    }

    /**
     * Run this class as a program
     * 
     * @param args
     *            The command line arguments.
     * 
     * @throws Exception
     *             Exception we don't recover from.
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("usage: ToHtml inputWorkbook outputXMLFile");
            return;
        }

        ToXML toHtml = create(args[0], new PrintWriter(new FileWriter(args[1])));
        toHtml.printPage();
    }

    public void printPage() throws IOException {
        try {
            ensureOut();
            out.format("<?xml version='1.0' encoding='iso-8859-1'?>\n"
                    + "<?xml-stylesheet type=\"text/xsl\" href=\"spreadsheet.xsl\"?>\n");
            printSheets();

        } finally {
            if (out != null)
                out.close();
            if (output instanceof Closeable) {
                Closeable closeable = (Closeable) output;
                closeable.close();
            }
        }
    }

    private void ensureOut() {
        if (out == null)
            out = new Formatter(output);
    }

    private void printSheets() {
        ensureOut();
        int total = wb.getNumberOfSheets();
        out.format("<spreadsheets>");
        for (int c = 0; c < total; c++) {
            currentSheet = c;
            Sheet sheet = wb.getSheetAt(c);
            printSheet(sheet);
            out.format("%n");
        }
        out.format("</spreadsheets>");
    }

    public void printSheet(Sheet sheet) {
        ensureOut();
        out.format("<Table name=\"%s\">%n", sheet.getSheetName());
        printSheetContent(sheet);
        out.format("</Table>%n");
    }

    private void ensureColumnBounds(Sheet sheet) {
        if (gotBounds)
            return;

        Iterator<Row> iter = sheet.rowIterator();
        firstColumn = (iter.hasNext() ? Integer.MAX_VALUE : 0);
        endColumn = 0;
        while (iter.hasNext()) {
            Row row = iter.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
        gotBounds = true;
    }

    private void printColumnHeads() {

        out.format("<ColumnHeaders>%n");
        colNumbers = new HashMap<Integer, String>(endColumn);
        StringBuilder colName = new StringBuilder();
        out.format("    <ColumnHeader>%s</ColumnHeader>%n", "RowID\\ColID");
        for (int i = firstColumn; i < endColumn; i++) {
            colName.setLength(0);
            int cnum = i;
            do {
                colName.insert(0, (char) ('A' + cnum % 26));
                cnum /= 26;
            } while (cnum > 0);
            colNumbers.put(cnum, colName.toString());

            out.format("    <ColumnHeader>%s</ColumnHeader>%n", colName);
        }

        out.format("</ColumnHeaders>%n");
    }

    private void printSheetContent(Sheet sheet) {
        ensureColumnBounds(sheet);
        printColumnHeads();

        cellsToFormula = new HashMap<String, List<String>>();
        cellToFormulaConverted = new HashMap<String, String>();
        crToParent = new HashMap<String, String>();
        
        FormulaParsingWorkbook fpwb;
        FormulaRenderingWorkbook frwb;
        if (xswb != null) {
            XSSFEvaluationWorkbook w = XSSFEvaluationWorkbook.create(xswb);
            frwb = w;
            fpwb = w;
        } else if (hswb != null) {
            HSSFEvaluationWorkbook w = HSSFEvaluationWorkbook.create(hswb);
            frwb = w;
            fpwb = w;
        }

        else
            return;
        // first we need to determine all the dependencies ofr each formula
        Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            Row row = rows.next();
            for (int i = firstColumn; i < endColumn; i++) {
                if (i >= row.getFirstCellNum() && i < row.getLastCellNum()) {
                    Cell cell = row.getCell(i);
                    if (cell != null) {
                        if (cell.getCellType() == Cell.CELL_TYPE_FORMULA)
                            try {
                                parseFormula(cell, fpwb, frwb);
                            } catch (Exception x) {

                            }
                    }
                }
            }
        }
        rows = sheet.rowIterator();

        while (rows.hasNext()) {
            Row row = rows.next();
            int rowNumber = row.getRowNum() + 1;
            out.format("  <TableRow>%n");
            out.format("    <RowHeader>%d</RowHeader>%n", rowNumber);
            out.format("  <TableCells>%n");
            for (int i = firstColumn; i < endColumn; i++) {
                String content = "0";
                String attrs = "";
                CellStyle style = null;
                String valueType = "float";
                Cell cell = row.getCell(i);
                CellReference c = new CellReference(rowNumber - 1, i);
                attrs += " cellID=\"." + c.formatAsString() + "\"";

                String cr = c.formatAsString();
                // if (i >= row.getFirstCellNum() && i < row.getLastCellNum()) {

                if (cell != null
                        && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    try {
                        attrs += " cellFormula=\"" + cell.getCellFormula()
                                + "\"";
                    } catch (Exception x) {
                        attrs += " cellFormula=\"FORMULA ERROR\"";
                    }
                } else {
                    List<String> cfrl = cellsToFormula.get(cr);
                    String formula = "";
                    HashSet<String> visited = new HashSet<String>();
                    if (cfrl != null)
                        for (String cfr : cfrl)
                            while (cfr != null) {
                                if (visited.contains(cfr)) {
                                    cfr = crToParent.get(cfr);
                                    continue;
                                } else
                                    visited.add(cfr);
                                formula += cellToFormulaConverted.get(cfr);
                                cfr = crToParent.get(cfr);
                                if (cfr != null)
                                    formula += " || ";
                            }
                    if (formula.length() > 0)
                        attrs += " formula=\"" + formula + "\"";
                }
                if (cell != null) {
                    style = cell.getCellStyle();
                    // Set the value that is rendered for the cell
                    // also applies the format

                    try {
                        CellFormat cf = CellFormat.getInstance(style
                                .getDataFormatString());
                        CellFormatResult result = cf.apply(cell);
                        content = result.text;
                    } catch (Exception x) {
                        content = "DATA FORMULA ERROR ";
                    }

                }
                // }
                attrs += " value_type=\"" + valueType + "\"";
                attrs += " value=\"" + content + "\"";
                out.format("    <TableCell  %s>%s</TableCell>%n", // class=%s
                                                                  // styleName(style),
                        attrs, content);
            }
            out.format(" </TableCells> </TableRow>%n%n");
        }
    }

    private void parseFormula(Cell cell, FormulaParsingWorkbook fpwb,
            FormulaRenderingWorkbook frwb) {
        CellReference c = new CellReference(cell);
        String cr = c.formatAsString();

        Ptg[] pp = FormulaParser.parse(cell.getCellFormula(), fpwb,
                FormulaType.CELL, currentSheet);

        for (Ptg p : pp) {
            if (p instanceof RefPtg) {
                RefPtg a = (RefPtg) p;
                Cell dest = cell.getSheet().getRow(a.getRow())
                        .getCell(a.getColumn());
                if (dest != null
                        && dest.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    String cr2 = new CellReference(dest).formatAsString();
                    crToParent.put(cr2, cr);
                }
                List<String> ls = cellsToFormula.get(a.toFormulaString());
                if (ls == null) {
                    ls = new LinkedList<String>();
                    ls.add(cr);
                    cellsToFormula.put(a.toFormulaString(), ls);
                } else
                    ls.add(cr);

            }
            if (p instanceof AreaPtg) {
                AreaPtg a = (AreaPtg) p;
                for (int i = a.getFirstColumn(); i <= a.getLastColumn(); i++) {
                    for (int k = a.getFirstRow(); k <= a.getLastRow(); k++) {
                        String cc = new CellReference(k, i).formatAsString();

                        Cell dest = cell.getSheet().getRow(k).getCell(i);
                        if (dest != null
                                && dest.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            String cr2 = new CellReference(dest)
                                    .formatAsString();
                            crToParent.put(cr2, cr);
                        }

                        List<String> ls = cellsToFormula.get(cc);
                        if (ls == null) {
                            ls = new LinkedList<String>();
                            ls.add(cr);
                            cellsToFormula.put(cc, ls);
                        } else
                            ls.add(cr);
                    }
                }
            }
        }

        String cellF = "[." + cr + "]="
                + FormulaRenderer.toFormulaString(frwb, pp);
        System.out.println(cellF);
        cellToFormulaConverted.put(cr, cellF);

    }

}