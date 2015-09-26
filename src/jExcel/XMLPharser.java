package jExcel;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.TimeZone;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

public class XMLPharser {

    private ArrayList<Cell> rawCells;
    private Cell[][] cells;
    private Book book;

    public Book getBook() {
        return book;
    }

    public XMLPharser(Book book, String path) throws XMLStreamException, FileNotFoundException {
        this.book = book;
        String tagContent = null;
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLStreamReader reader;
        reader = factory.createXMLStreamReader(new FileInputStream(path));

        int i = 0;
        int j = 0;
        int sheetIdx = -1;
        String formula = "";
        rawCells = new ArrayList();
        while (reader.hasNext()) {
            int event = reader.next();
            switch (event) {
                case XMLStreamConstants.START_ELEMENT:
                    if ("Row".equals(reader.getLocalName())) {
                        if (reader.getAttributeValue(null, "Index") == null) {
                            i++;
                        } else {
                            i = Integer.parseInt(reader.getAttributeValue(null, "Index"));
                        }

                    } else if ("Cell".equals(reader.getLocalName())) {
                        if (reader.getAttributeValue(null, "Index") == null) {
                            j++;
                        } else {
                            j = Integer.parseInt(reader.getAttributeValue(null, "Index"));
                        }
                        formula = reader.getAttributeValue(null, "Formula") == null ? "" : reader.getAttributeValue(null, "Formula").replaceFirst("=", "");

                    } else if ("Table".equals(reader.getLocalName())) {
                        int cRow = Integer.parseInt(reader.getAttributeValue(null, "ExpandedRowCount") == null ? "0" : reader.getAttributeValue(null, "ExpandedRowCount"));
                        int cCount = Integer.parseInt(reader.getAttributeValue(null, "ExpandedColumnCount") == null ? "0" : reader.getAttributeValue(null, "ExpandedColumnCount"));

                    }
                    if ("Worksheet".equals(reader.getLocalName())) {
                        sheetIdx++;

                        if (sheetIdx == book.getSheets().toArray().length) {
                            book.addNewSheet();
                            while (sheetIdx == book.getSheets().toArray().length) {/* wait..*/

                            }
                            cells = book.getSheets().get(sheetIdx).getSheetModel().cells;
                        }
                    }

                    break;

                case XMLStreamConstants.CHARACTERS:
                    tagContent = reader.getText().trim();
                    break;

                case XMLStreamConstants.END_ELEMENT:
                    switch (reader.getLocalName()) {
                        case "Data":
                            cells[i - 1][j - 1] = new Cell(book.getSheets().get(sheetIdx), i - 1, j - 1, tagContent, formula);
                            if (formula.length() > 1) {
                                rawCells.add(cells[i - 1][j - 1]);
                            }
                            break;
                        case "Row":
                            j = 0;
                            break;
                        case "Table":
                            i = 0;
                            break;
                        case "Workbook":
                            for (Cell cell : rawCells) {
                                cell.Compute("=" + cell.getFormula(), cell.getSheet().getBook());
                            }

                    }
                    break;

                case XMLStreamConstants.START_DOCUMENT:
                    break;
            }

        }

    }

    public static void createXmlFile(Book book, String path) throws IOException {
        String timeStamp=String.format("%tFT%<tRZ",Calendar.getInstance(TimeZone.getTimeZone("Z")));
        String xml = "<?xml version=\"1.0\"?>\n"
                + "<?mso-application progid=\"Excel.Sheet\"?>\n"
                + "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n"
                + " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n"
                + " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n"
                + " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n"
                + " xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n"
                + " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n"
                + "  <Author>adm.trine@gmail.com</Author>\n"
                + "  <LastAuthor>adm.trine@gmail.com</LastAuthor>\n"
                + "  <Created>"+timeStamp+"</Created>\n"
                + "  <LastSaved>"+timeStamp+"</LastSaved>\n"
                + "  <Version>15.00</Version>\n"
                + " </DocumentProperties>\n"
                + " <OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\">\n"
                + "  <AllowPNG/>\n"
                + " </OfficeDocumentSettings>\n"
                + " <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">\n"
                + "  <WindowHeight>7425</WindowHeight>\n"
                + "  <WindowWidth>20490</WindowWidth>\n"
                + "  <WindowTopX>0</WindowTopX>\n"
                + "  <WindowTopY>0</WindowTopY>\n"
                + "  <ActiveSheet>0</ActiveSheet>\n"
                + "  <MaxChange>0.0001</MaxChange>\n"
                + "  <ProtectStructure>False</ProtectStructure>\n"
                + "  <ProtectWindows>False</ProtectWindows>\n"
                + " </ExcelWorkbook>\n"
                + " <Styles>\n"
                + "  <Style ss:ID=\"Default\" ss:Name=\"Normal\">\n"
                + "   <Alignment ss:Vertical=\"Bottom\"/>\n"
                + "   <Borders/>\n"
                + "   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>\n"
                + "   <Interior/>\n"
                + "   <NumberFormat/>\n"
                + "   <Protection/>\n"
                + "  </Style>\n"
                + " </Styles>";
        for (Sheet sheet : book.getSheets()) {
            xml += "<Worksheet ss:Name=\"" + sheet.getSheetName() + "\"><Table ss:ExpandedColumnCount=\"50\" ss:ExpandedRowCount=\"50\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\">";
            int ri = 1;
            for (Cell[] row : sheet.getSheetModel().cells) {
                String cellxml = "";
                int ci = 0;
                for (Cell cell : row) {
                    ci++;
                    if (cell.getValue().equals("") && cell.getRformula().equals("")) {
                        continue;
                    }
                    cellxml += "<Cell ss:Index=\"" + ci + "\" ss:Formula=\"" + cell.getRformula() + "\"><Data ss:Type=\"Number\">" + cell.getValue() + "</Data></Cell>\n";
                }
                if (cellxml.length() > 1) {
                    xml += "<Row ss:Index=\"" + ri + "\">\n" + cellxml + "</Row>\n";
                }
                ri++;
            }
            xml += "</Table>\n"
                    + "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n"
                    + "   <PageSetup>\n"
                    + "    <Header x:Margin=\"0.3\"/>\n"
                    + "    <Footer x:Margin=\"0.3\"/>\n"
                    + "    <PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/>\n"
                    + "   </PageSetup>\n"
                    + "   <Selected/>\n"
                    + "   <Panes>\n"
                    + "    <Pane>\n"
                    + "     <Number>3</Number>\n"
                    + "     <ActiveRow>0</ActiveRow>\n"
                    + "    </Pane>\n"
                    + "   </Panes>\n"
                    + "   <ProtectObjects>False</ProtectObjects>\n"
                    + "   <ProtectScenarios>False</ProtectScenarios>\n"
                    + "  </WorksheetOptions>\n"
                    + " </Worksheet>";
        }
        xml += "</Workbook>";
        System.out.println("" + xml);
        //out.print("nikan");
        PrintWriter out = new PrintWriter(new BufferedWriter(new FileWriter(path)));
        out.write(xml);
        out.close();
    }

    /**
     * @return the rawCells
     */
    public ArrayList<Cell> getRawCells() {
        return rawCells;
    }

    /**
     * @param rawCells the rawCells to set
     */
    public void setRawCells(ArrayList<Cell> rawCells) {
        this.rawCells = rawCells;
    }

    /**
     * @return the cells
     */
    public Cell[][] getCells() {
        return cells;
    }

    /**
     * @param cells the cells to set
     */
    public void setCells(Cell[][] cells) {
        this.cells = cells;
    }

    /**
     * @param book the book to set
     */
    public void setBook(Book book) {
        this.book = book;
    }
}
