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
/**
 * This class will analyze and report the specific features used in an Excel document. 
 * This is intended as an exemplar for the feature based preservation analysis, described in the SHAMAN deliverable 9.3
 * and Deliverable 4.4
 * 
 * @author Fabio Corubolo
 */

package uk.ac.liverpool.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFObjectData;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.PaneInformation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.Namespace;
import org.jdom.output.Format;
import org.jdom.output.XMLOutputter;

/**
 * Excel feature analysis. This class allows analysing the main features present
 * in an Excel (both binary and XML based) file. This will help classifying the
 * file and planning an appropriate preservation and long term access strategy.
 * 
 * 
 * @author Fabio Corubolo
 * 
 */

public class ExcelFeatureAnalysis {

    public static final String VERSION = "0.3";
    public static final String SW_ID = ExcelFeatureAnalysis.class
    .getCanonicalName();
    private static final String MIME = "application/vnd.ms-excel";
    // The feature analysis namespace
    public static final Namespace fa = Namespace
    .getNamespace("http://shaman.disi.unitn.it/FeatureAnalysis");
    // Dublin core
    public static final Namespace dc = Namespace.getNamespace("dc",
    "http://purl.org/dc/elements/1.1/");
    // The data format specific part (in this case, spreadsheets)
    public static final Namespace sn = Namespace.getNamespace("ss",
    "http://shaman.disi.unitn.it/Spreadsheet");

    Workbook wb;
    HSSFWorkbook hswb;
    XSSFWorkbook xswb;
    File f;

    /**
     * 
     * @param w
     * @param f
     */

    private ExcelFeatureAnalysis(Workbook w, File f) {
        wb = w;
        if (wb instanceof HSSFWorkbook) {
            hswb = (HSSFWorkbook) wb;
        } else if (wb instanceof XSSFWorkbook) {
            xswb = (XSSFWorkbook) wb;
        }
        this.f = f;

    }

    /**
     * 
     * This method will instantiate the workbook and return the feature analysis
     * object.
     * 
     * @param in
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static ExcelFeatureAnalysis create(File in) throws IOException,
    InvalidFormatException, EncryptedDocumentException {
        Workbook wb = WorkbookFactory.create(new FileInputStream(in));
        return new ExcelFeatureAnalysis(wb, in);

    }

    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            System.err.println("usage: ToHtml inputWorkbook");

            return;
        }

        if (args.length > 2 && args[2].equals("-something"))
            ;
        String s = analyse(new File(args[0]));
        System.out.println(s);
    }

    /**
     * 
     * The main analysis method. The returned string will be a properly formed
     * XML enlisting the file features.
     * 
     * @return XML string listing the features used in the file.
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static String analyse(File in) throws InvalidFormatException,
    IOException {

        ExcelFeatureAnalysis efa;
        try {
            efa  = create(in); 
        }
        catch (EncryptedDocumentException e) {
            Element r = new Element("featureanalysis", fa);
            SimpleDateFormat sdt = new SimpleDateFormat(
            "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
            r.addContent(new Element("processing_date", fa).addContent(sdt
                    .format(new Date())));
            r.addContent(new Element("processingSoftware", fa).setAttribute("name",
                    SW_ID).setAttribute("version", VERSION));
            Element da = new Element("object", fa);
            r.addContent(da);
            da.setAttribute("filename", in.getName());
            da.setAttribute("lastModified", sdt.format(new Date(in.lastModified())));
            da.setAttribute("mimeType", MIME);
            da.addContent(new Element("encrypted", fa));
            Document d = new Document();
            d.setRootElement(r);
            XMLOutputter o = new XMLOutputter(Format.getPrettyFormat());
            String res = o.outputString(d);

            return res;
        }
        
        Element r = new Element("featureanalysis", fa);
        SimpleDateFormat sdt = new SimpleDateFormat(
        "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
        r.addContent(new Element("processing_date", fa).addContent(sdt
                .format(new Date())));
        r.addContent(new Element("processingSoftware", fa).setAttribute("name",
                SW_ID).setAttribute("version", VERSION));
        Element da = new Element("object", fa);
        r.addContent(da);
        da.setAttribute("filename", in.getName());
        da.setAttribute("lastModified", sdt.format(new Date(in.lastModified())));
        da.setAttribute("mimeType", MIME);

        // end of the generic part, beginning of the file specific
        analyseSpreadsheet(da, efa);

        // finishing up, formatting and return string
        Document d = new Document();
        d.setRootElement(r);
        XMLOutputter o = new XMLOutputter(Format.getPrettyFormat());
        String res = o.outputString(d);

        return res;
    }

    // Analysis at the file level
    private static void analyseSpreadsheet(Element da, ExcelFeatureAnalysis efa) {

        Element s = new Element("spreadsheets", sn);
        da.addContent(s);
        s.setAttribute("numberOfSheets", "" + efa.wb.getNumberOfSheets());
        // workbook wide features

        List<? extends PictureData> allPictures = efa.wb.getAllPictures();
        if (allPictures != null && allPictures.size() > 0) {
            Element oo = new Element("Pictures", sn);
            s.addContent(oo);
            for (PictureData pd : allPictures) {
                Element ob = new Element("Picture", sn);
                ob.setAttribute("mimeType", pd.getMimeType());
                oo.addContent(ob);
            }
        }

        int numfonts = efa.wb.getNumberOfFonts();
        if (numfonts > 0) {
            Element oo = new Element("Fonts", sn);
            s.addContent(oo);
            for (int i = 0; i < numfonts; i++) {
                Font cs = efa.wb.getFontAt((short) i);
                Element ob = new Element("Font", sn);
                ob.setAttribute("Name", cs.getFontName());

                ob.setAttribute("Charset", "" + cs.getCharSet());
                oo.addContent(ob);
            }
        }

        if (efa.hswb != null) {

            DocumentSummaryInformation dsi = efa.hswb
            .getDocumentSummaryInformation();
            if (dsi != null)
                s.setAttribute("OSVersion", "" + dsi.getOSVersion());
            // Property[] properties = dsi.getProperties();
            // CustomProperties customProperties = dsi.getCustomProperties();

            List<HSSFObjectData> eo = efa.hswb.getAllEmbeddedObjects();
            if (eo != null && eo.size() > 0) {
                Element oo = new Element("EmbeddedObjects", sn);
                s.addContent(oo);
                for (HSSFObjectData o : eo) {
                    Element ob = new Element("EmbeddedObject", sn);
                    ob.setAttribute("name", o.getOLE2ClassName());
                    oo.addContent(ob);
                }

            }
        } else if (efa.xswb != null) {
            try {
                POIXMLProperties properties = efa.xswb.getProperties();
                List<PackagePart> allEmbedds = efa.xswb.getAllEmbedds();
                if (allEmbedds != null && allEmbedds.size() > 0) {
                    Element oo = new Element("EmbeddedObjects", sn);
                    s.addContent(oo);

                    for (PackagePart p : allEmbedds) {
                        Element ob = new Element("EmbeddedObject", sn);
                        ob.setAttribute("mimeType", p.getContentType());
                        ob.setAttribute("name", p.getPartName().getName());

                        oo.addContent(ob);
                    }
                }
            } catch (OpenXML4JException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }

        }
        int nn = efa.wb.getNumberOfNames();
        if (nn > 0) {
            Element oo = new Element("NamedCells", sn);
            s.addContent(oo);
        }

        // sheet specific features
        int total = efa.wb.getNumberOfSheets();
        for (int c = 0; c < total; c++) {
            Sheet sheet = efa.wb.getSheetAt(c);
            Element single = new Element("sheet", sn);
            s.addContent(single);
            analyseSheet(sheet, single, sn, efa);
        }
    }

    // Analysis at the sheet level

    private static void analyseSheet(Sheet ss, Element s, Namespace n,
            ExcelFeatureAnalysis efa) {
        // generic part
        boolean costumFormatting = false;
        boolean formulae = false;
        boolean UDF = false;
        boolean hasComments = false;

        Set<String> udfs = new HashSet<String>();
        FormulaEvaluator evaluator = ss.getWorkbook().getCreationHelper()
        .createFormulaEvaluator();

        s.setAttribute("name", ss.getSheetName());
        s.setAttribute("firstRow", "" + ss.getFirstRowNum());
        s.setAttribute("lastRow", "" + ss.getLastRowNum());
        try {
            s.setAttribute("forceFormulaRecalc",
                    "" + ss.getForceFormulaRecalculation());
        } catch (Throwable x) {
            //x.printStackTrace();
        } 

        // shapes in detail? 
        Footer footer = ss.getFooter();
        if (footer != null) {
            s.setAttribute("footer", "true");
        }
        Header header = ss.getHeader();
        if (header != null) {
            s.setAttribute("header", "true");
        }
        PaneInformation paneInformation = ss.getPaneInformation();
        if (paneInformation != null) {
            s.setAttribute("panels", "true");
        }

        HSSFSheet hs = null;
        XSSFSheet xs = null;
        if (ss instanceof HSSFSheet) {
            hs = (HSSFSheet) ss;
            try {
                if (hs.getDrawingPatriarch() != null) {
                    if (hs.getDrawingPatriarch().containsChart())
                        s.setAttribute("charts", "true");
                    if (hs.getDrawingPatriarch().countOfAllChildren() > 0)
                        s.setAttribute("shapes", "true");
                }
            } catch (Exception x) {
                x.printStackTrace();
            }

            if (hs.getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0) {
                s.setAttribute("conditionalFormatting", "true");
            }
        }
        if (ss instanceof XSSFSheet) {
            xs = (XSSFSheet) ss;

        }
        Iterator<Row> rows = ss.rowIterator();

        int firstColumn = (rows.hasNext() ? Integer.MAX_VALUE : 0);
        int endColumn = 0;
        while (rows.hasNext()) {
            Row row = rows.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
        s.setAttribute("firstColumn", "" + firstColumn);
        s.setAttribute("lastColumn", "" + endColumn);
        rows = ss.rowIterator();
        while (rows.hasNext()) {
            Row row = rows.next();
            for (Cell cell : row)
                if (cell != null) {
                    try {
                        if (!cell.getCellStyle().getDataFormatString()
                                .equals("GENERAL"))
                            costumFormatting = true;
                    } catch (Throwable t) {}

                    if (cell.getCellComment()!=null)
                        hasComments = true;
                    switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        // System.out.println(cell.getRichStringCellValue().getString());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        //                        if (DateUtil.isCellDateFormatted(cell)) {
                        //                            // System.out.println(cell.getDateCellValue());
                        //                        } else {
                        //                            // System.out.println(cell.getNumericCellValue());
                        //                        }
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        // System.out.println(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        // System.out.println(cell.getCellFormula());
                        formulae = true;
                        if (!UDF)
                            try {
                                evaluator.evaluate(cell);
                            } catch (Exception x) {
                                if (x instanceof NotImplementedException) {
                                    Throwable e = x;

                                    //e.printStackTrace();
                                    while (e!=null) {
                                        for (StackTraceElement c : e.getStackTrace()) {
                                            if (c.getClassName().contains(
                                            "UserDefinedFunction")) {
                                                UDF = true;
                                                //System.out.println("UDF " + e.getMessage());
                                                udfs.add(e.getMessage());
                                            }
                                        }
                                        e = e.getCause();
                                    }

                                }
                            }
                            break;
                    default:
                    }

                }
        }
        if (costumFormatting) {
            Element cf = new Element("customisedFormatting", sn);
            s.addContent(cf);
        }
        if (formulae) {
            Element cf = new Element("formulae", sn);
            s.addContent(cf);
        }
        if (UDF) {
            Element cf = new Element("userDefinedFunctions", sn);
            for (String sss: udfs)
                cf.addContent(new Element("userDefinedFunction",sn).setAttribute("functionName",sss));
            s.addContent(cf);
        }
        if (hasComments) {
            Element cf = new Element("cellComments", sn);
            s.addContent(cf);
        }
    }
}
