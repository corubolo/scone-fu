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
package uk.ac.liverpool.spreadsheet;

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
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This example shows how to export a spreadsheet in XML using the classes for
 * spreadsheet display. This includes formulae transformation .
 * based on toHtml from Ken Arnolds
 * 
 * @author Fabio Corubolo, University of Trento
 */


public class ToXML {
    private final Workbook wb;
    private Appendable output;
    private Formatter out;
    private boolean gotBounds;
    private int firstColumn;
    private int endColumn;
    private int currentSheet;
    private HSSFWorkbook hswb;
    private XSSFWorkbook xswb;

    private boolean evaluateFormulae = false;

    public boolean isEvaluateFormulae() {
        return evaluateFormulae;
    }

    public void setEvaluateFormulae(boolean evaluateFormulae) {
        this.evaluateFormulae = evaluateFormulae;
    }

    private Map<Integer, String> colNumbers;

    // every cell (value) can be referred to by multiple formulas
    private Map<String, List<String>> cellsToFormula;
    // every cell (formula) can be referred to by multiple formulas
    private Map<String,List<String>> crToParent;

    private Map<String, String> cellToFormulaConverted;


    public static ToXML create(InputStream in)
    throws IOException, InvalidFormatException {

        Workbook wb = WorkbookFactory.create(in);
        return new ToXML(wb);

    }

    private ToXML(Workbook wb) {
        if (wb == null)
            throw new NullPointerException("wb");
        this.wb = wb;

        if (wb instanceof HSSFWorkbook) {
            hswb = (HSSFWorkbook) wb;
        } else if (wb instanceof XSSFWorkbook) {
            xswb = (XSSFWorkbook) wb;
        }
    }

    /**
     * Run this class as a program
     * 
     * the Output file will be named inputWorkbook.[sheetNumber].xml
     * @param args
     *            The command line arguments.
     * 
     * @throws Exception
     *             Exception we don't recover from.
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            System.err
            .println("usage: ToXml inputWorkbook [-evaluate]/n the Output file will be named inputWorkbook.[sheetNumber].xml");

            return;
        }

        ToXML toHtml = create(new FileInputStream(args[0]));
        if (args.length > 1 && args[1].equals("-evaluate"))
            toHtml.evaluateFormulae = true;
        toHtml.convert(args[1]);
    }

    /** 
     * Spread sheet level conversion
     * the Output file will be named filename.[sheetNumber].xml
     * @param filename to convert 
     * @throws IOException
     */
    
    public void convert(String filename) throws IOException {
        if (evaluateFormulae) {
            FormulaEvaluator evaluator = wb.getCreationHelper()
            .createFormulaEvaluator();
            for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = wb.getSheetAt(sheetNum);
                for (Row r : sheet) {
                    for (Cell c : r) {
                        if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            evaluator.evaluateFormulaCell(c);
                        }
                    }
                }
            }
        }


        convertSheets(filename);


    }

    /** 
     * Spread sheet level conversion
     * the Output file will be named inputWorkbook.[sheetNumber].xml
     * @param File to convert
     * @throws IOException
     */
    private void convertSheets(String filename) throws IOException {
        int total = wb.getNumberOfSheets();
        
        String start = filename.substring(0,filename.lastIndexOf('.'));
        String end = filename.substring(filename.lastIndexOf('.'));

        for (int c = 0; c < total; c++) {
            try {
                Sheet sheet = wb.getSheetAt(c);
                if (!sheet.rowIterator().hasNext())
                    continue;

                output = new PrintWriter(new FileWriter(start+c+end));

                out = new Formatter(output);
                out.format("<?xml version='1.0' encoding='iso-8859-1'?>\n"
                        + "<?xml-stylesheet type=\"text/xsl\" href=\"spreadsheet.xsl\"?>\n");
                out.format("<spreadsheets>");
                currentSheet = c;

                printSheet(sheet);
                out.format("%n");
                out.format("</spreadsheets>");
            } finally {
                if (out != null)
                    out.close();
                if (output instanceof Closeable) {
                    Closeable closeable = (Closeable) output;
                    closeable.close();
                }
            }
        }

    }

    private void printSheet(Sheet sheet) {
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
        crToParent = new HashMap<String, List<String>>();
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
                    attrs += " readOnly=\"readOnly\""; 
                    try {
                        attrs += " cellFormula=\"" + cell.getCellFormula()
                        + "\"";
                    } catch (Exception x) {
                        attrs += " cellFormula=\"FORMULA ERROR\"";
                    }
                } else {
                    List<String> cfrl = cellsToFormula.get(cr);
                    StringBuffer formula = new StringBuffer("");

                    if (cfrl != null) {
                        List<String>refs = new LinkedList<String>();
                        visit(cfrl, refs);
                        System.out.println(refs);
                        cleanup(refs);
                        for (String s:refs) {
                            formula.append(cellToFormulaConverted.get(s));
                            formula .append(" || ");
                        }
                    }
                    if (formula.length() > 0)
                        attrs += " formula=\"" + formula.substring(0,formula.length()-4) + "\"";
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

    /**
     * This method will remove all duplicates starting from the right of the list
     * 
     * @param refs
     */
    private void cleanup(List<String> refs) {
        HashSet<String> visited = new HashSet<String>();
        for (int i=refs.size()-1;i>=0;i--) {
            String h = refs.get(i);
            if (!visited.add(h))
                refs.remove(i);
        }

    }

    private void visit(List<String> cfrl,List<String> refs) {
        List<String> parents = new LinkedList<String>();

        for (String cfr : cfrl) {
            if (cfr == null)
                continue;
            List<String> list = crToParent.get(cfr);
            if (list!=null)
                parents.addAll(list);
            refs.add(cfr);
        }
        if (parents.size()>0)
            visit(parents, refs);
        return ;
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
                    List<String> ls = crToParent.get(cr2);
                    if (ls == null) {
                        ls = new LinkedList<String>();
                        crToParent.put(cr2, ls);
                    }
                    ls.add(cr);
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
                            List<String> ls = crToParent.get(cr2);
                            if (ls == null) {
                                ls = new LinkedList<String>();
                                crToParent.put(cr2, ls);
                            }
                            ls.add(cr);
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