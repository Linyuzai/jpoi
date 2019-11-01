package com.github.linyuzai.jpoi.excel.read.sax;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.ZipPackagePart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.*;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Pattern;

public class SaxWorkbook implements Workbook {

    private List<Sheet> sheets = new ArrayList<>();
    private SaxSheet tempSheet;

    private SaxCreationHelper creationHelper = new SaxCreationHelper();

    public SaxWorkbook(InputStream is) throws IOException {
        try {
            OPCPackage pkg = OPCPackage.open(is);
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable st = xssfReader.getStylesTable();
            SharedStringsTable sst = xssfReader.getSharedStringsTable();
            XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
            //parser.setContentHandler(new SaxHandler());
            parser.setContentHandler(new SaxHandler(st, sst, false));
            //XSSFSheetXMLHandler.SheetContentsHandler
            XSSFReader.SheetIterator xsi = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            List<PackagePart> packageParts = pkg.getPartsByName(Pattern.compile("/xl/media/.*?"));
            //packageParts.get(0).getInputStream()
            //Map<String, >
            for (PackagePart packagePart : packageParts) {
                //packagePart.getRelationship("")
                packagePart.getPartName().getName();
                //((ZipPackagePart)packagePart).getZipArchive().get
            }
            while (xsi.hasNext()) { //遍历sheet
                tempSheet = new SaxSheet();
                sheets.add(tempSheet);
                //curRow = 1; //标记初始行为第一行
                //sheetIndex++;
                InputStream sheet = xsi.next(); //sheets.next()和sheets.getSheetName()不能换位置，否则sheetName报错
                tempSheet.setName(xsi.getSheetName());
                List<XSSFShape> shapes = xsi.getShapes();
                CommentsTable commentsTable = xsi.getSheetComments();
                //xsi.getSheetPart().getRelationships()
                for (XSSFShape shape : shapes) {
                    //shape.getDrawing().getPackagePart().getRelationships().getRelationship(0).getId();
                    //shape.getDrawing().addRelation()
                    //String blipId = shape.getBlipFill().getBlip().getEmbed();
                    if (shape instanceof XSSFPicture) {
                        String blipId = ((XSSFPicture) shape).getCTPicture().getBlipFill().getBlip().getEmbed();
                        PackageRelationship packageRelationship = shape.getDrawing().getPackagePart().getRelationships().getRelationshipByID(blipId);
                        packageRelationship.getTargetURI();
                        //XSSFPicture
                        PictureData pictureData = ((Picture) shape).getPictureData();
                        System.out.println(pictureData);
                    }
                }
                tempSheet.setDrawing(new SaxDrawing(xsi.getShapes()));
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource); //解析excel的每条记录，在这个过程中startElement()、characters()、endElement()这三个函数会依次执行
                sheet.close();
            }
            is.close();
        } catch (OpenXML4JException | SAXException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public int getActiveSheetIndex() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setActiveSheet(int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getFirstVisibleTab() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setFirstVisibleTab(int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetOrder(String sheetName, int pos) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSelectedTab(int index) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetName(int sheet, String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getSheetName(int sheet) {
        return ((SaxSheet) sheets.get(sheet)).getName();
    }

    @Override
    public int getSheetIndex(String name) {
        for (int i = 0; i < sheets.size(); i++) {
            if (((SaxSheet) sheets.get(i)).getName().equals(name)) {
                return i;
            }
        }
        return -1;
    }

    @Override
    public int getSheetIndex(Sheet sheet) {
        for (int i = 0; i < sheets.size(); i++) {
            if (sheets.get(i) == sheet) {
                return i;
            }
        }
        return -1;
    }

    @Override
    public Sheet createSheet() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Sheet createSheet(String sheetName) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Sheet cloneSheet(int sheetNum) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Sheet> sheetIterator() {
        return sheets.iterator();
    }

    @Override
    public int getNumberOfSheets() {
        return sheets.size();
    }

    @Override
    public Sheet getSheetAt(int index) {
        return sheets.get(index);
    }

    @Override
    public Sheet getSheet(String name) {
        for (Sheet sheet : sheets) {
            if (((SaxSheet) sheet).getName().equals(name)) {
                return sheet;
            }
        }
        return null;
    }

    @Override
    public void removeSheetAt(int index) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Font createFont() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Font findFont(boolean bold, short color, short fontHeight, String name, boolean italic, boolean strikeout, short typeOffset, byte underline) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getNumberOfFonts() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getNumberOfFontsAsInt() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Font getFontAt(short idx) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Font getFontAt(int idx) {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellStyle createCellStyle() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getNumCellStyles() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellStyle getCellStyleAt(int idx) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void write(OutputStream stream) throws IOException {
        throw new UnsupportedOperationException();
    }

    @Override
    public void close() throws IOException {

    }

    @Override
    public int getNumberOfNames() {
        return sheets.size();
    }

    @Override
    public Name getName(String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends Name> getNames(String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends Name> getAllNames() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Name getNameAt(int nameIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Name createName() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getNameIndex(String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(int index) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(Name name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int linkExternalWorkbook(String name, Workbook workbook) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setPrintArea(int sheetIndex, String reference) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getPrintArea(int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removePrintArea(int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Row.MissingCellPolicy getMissingCellPolicy() {
        return Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;
    }

    @Override
    public void setMissingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {
        throw new UnsupportedOperationException();
    }

    @Override
    public DataFormat createDataFormat() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int addPicture(byte[] pictureData, int format) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends PictureData> getAllPictures() {
        List<PictureData> pictureData = new ArrayList<>();
        for (Sheet sheet : sheets) {
            Drawing<?> drawing = sheet.getDrawingPatriarch();
            if (drawing != null) {
                for (Shape shape : drawing) {
                    if (shape instanceof Picture) {
                        pictureData.add(((Picture) shape).getPictureData());
                    }
                }
            }
        }
        return pictureData;
    }

    @Override
    public CreationHelper getCreationHelper() {
        return creationHelper;
    }

    @Override
    public boolean isHidden() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHidden(boolean hiddenFlag) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isSheetHidden(int sheetIx) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isSheetVeryHidden(int sheetIx) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetHidden(int sheetIx, boolean hidden) {
        throw new UnsupportedOperationException();
    }

    @Override
    public SheetVisibility getSheetVisibility(int sheetIx) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetVisibility(int sheetIx, SheetVisibility visibility) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void addToolPack(UDFFinder toolPack) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setForceFormulaRecalculation(boolean value) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getForceFormulaRecalculation() {
        throw new UnsupportedOperationException();
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int addOlePackage(byte[] oleData, String label, String fileName, String command) throws IOException {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Sheet> iterator() {
        return sheets.iterator();
    }

    /**
     * These are the different kinds of cells we support.
     * We keep track of the current one between
     * the start and end.
     */
    enum XssfDataType {
        BOOLEAN,
        ERROR,
        FORMULA,
        INLINE_STRING,
        SST_STRING,
        NUMBER,
    }

    private enum EmptyCellCommentsCheckType {
        CELL,
        END_OF_ROW,
        END_OF_SHEET_DATA
    }

    public class SaxHandler extends DefaultHandler {
        private final POILogger logger = POILogFactory.getLogger(SaxHandler.class);

        /**
         * Table with the styles used for formatting
         */
        private Styles stylesTable;

        /**
         * Table with cell comments
         */
        private Comments comments;

        /**
         * Read only access to the shared strings table, for looking
         * up (most) string cell's contents
         */
        private SharedStrings sharedStringsTable;

        /**
         * Where our text is going
         */
        //private final org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler output;

        // Set when V start element is seen
        private boolean vIsOpen;
        // Set when F start element is seen
        private boolean fIsOpen;
        // Set when an Inline String "is" is seen
        private boolean isIsOpen;
        // Set when a header/footer element is seen
        private boolean hfIsOpen;

        // Set when cell start element is seen;
        // used when cell close element is seen.
        private XssfDataType nextDataType;

        // Used to format numeric cell values.
        private short formatIndex;
        private String formatString;
        private final DataFormatter formatter;
        private int rowNum;
        private int nextRowNum;      // some sheets do not have rowNums, Excel can read them so we should try to handle them correctly as well
        private String cellRef;
        private boolean formulasNotResults;

        // Gathers characters as they are seen.
        private StringBuilder value = new StringBuilder(64);
        private StringBuilder formula = new StringBuilder(64);
        private StringBuilder headerFooter = new StringBuilder(64);

        private Queue<CellAddress> commentCellRefs;

        private SaxRow tempRow;

        /**
         * Accepts objects needed while parsing.
         *
         * @param styles  Table of styles
         * @param strings Table of shared strings
         */
        public SaxHandler(Styles styles, Comments comments, SharedStrings strings, DataFormatter dataFormatter, boolean formulasNotResults) {
            this.stylesTable = styles;
            this.comments = comments;
            this.sharedStringsTable = strings;
            //this.output = sheetContentsHandler;
            this.formulasNotResults = formulasNotResults;
            this.nextDataType = XssfDataType.NUMBER;
            this.formatter = dataFormatter;
            init(comments);
        }

        /**
         * Accepts objects needed while parsing.
         *
         * @param styles  Table of styles
         * @param strings Table of shared strings
         */
        public SaxHandler(Styles styles, SharedStrings strings, DataFormatter dataFormatter, boolean formulasNotResults) {
            this(styles, null, strings, dataFormatter, formulasNotResults);
        }

        /**
         * Accepts objects needed while parsing.
         *
         * @param styles  Table of styles
         * @param strings Table of shared strings
         */
        public SaxHandler(Styles styles, SharedStrings strings, boolean formulasNotResults) {
            this(styles, strings, new DataFormatter(), formulasNotResults);
        }

        private void init(Comments commentsTable) {
            if (commentsTable != null) {
                commentCellRefs = new LinkedList<>();
                for (Iterator<CellAddress> iter = commentsTable.getCellAddresses(); iter.hasNext(); ) {
                    commentCellRefs.add(iter.next());
                }
            }
        }

        private boolean isTextTag(String name) {
            if ("v".equals(name)) {
                // Easy, normal v text tag
                return true;
            }
            if ("inlineStr".equals(name)) {
                // Easy inline string
                return true;
            }
            if ("t".equals(name) && isIsOpen) {
                // Inline string <is><t>...</t></is> pair
                return true;
            }
            // It isn't a text tag
            return false;
        }

        @Override
        @SuppressWarnings("unused")
        public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {

            if (uri != null && !uri.equals(XSSFRelation.NS_SPREADSHEETML)) {
                return;
            }

            if (isTextTag(localName)) {
                vIsOpen = true;
                // Clear contents cache
                value.setLength(0);
            } else if ("is".equals(localName)) {
                // Inline string outer tag
                isIsOpen = true;
            } else if ("f".equals(localName)) {
                // Clear contents cache
                formula.setLength(0);

                // Mark us as being a formula if not already
                if (nextDataType == XssfDataType.NUMBER) {
                    nextDataType = XssfDataType.FORMULA;
                }

                // Decide where to get the formula string from
                String type = attributes.getValue("t");
                if (type != null && type.equals("shared")) {
                    // Is it the one that defines the shared, or uses it?
                    String ref = attributes.getValue("ref");
                    String si = attributes.getValue("si");

                    if (ref != null) {
                        // This one defines it
                        // TODO Save it somewhere
                        fIsOpen = true;
                    } else {
                        // This one uses a shared formula
                        // TODO Retrieve the shared formula and tweak it to
                        //  match the current cell
                        if (formulasNotResults) {
                            logger.log(POILogger.WARN, "shared formulas not yet supported!");
                        } /*else {
                   // It's a shared formula, so we can't get at the formula string yet
                   // However, they don't care about the formula string, so that's ok!
                }*/
                    }
                } else {
                    fIsOpen = true;
                }
            } else if ("oddHeader".equals(localName) || "evenHeader".equals(localName) ||
                    "firstHeader".equals(localName) || "firstFooter".equals(localName) ||
                    "oddFooter".equals(localName) || "evenFooter".equals(localName)) {
                hfIsOpen = true;
                // Clear contents cache
                headerFooter.setLength(0);
            } else if ("row".equals(localName)) {
                String rowNumStr = attributes.getValue("r");
                if (rowNumStr != null) {
                    rowNum = Integer.parseInt(rowNumStr) - 1;
                } else {
                    rowNum = nextRowNum;
                }
                //TODO output.startRow(rowNum);
                tempRow = (SaxRow) tempSheet.createRow(rowNum);//new SaxRow();
                //tempRow.setRowNum(rowNum);
                tempSheet.getRows().add(tempRow);
            }
            // c => cell
            else if ("c".equals(localName)) {
                // Set up defaults.
                this.nextDataType = XssfDataType.NUMBER;
                this.formatIndex = -1;
                this.formatString = null;
                cellRef = attributes.getValue("r");
                String cellType = attributes.getValue("t");
                String cellStyleStr = attributes.getValue("s");
                if ("b".equals(cellType))
                    nextDataType = XssfDataType.BOOLEAN;
                else if ("e".equals(cellType))
                    nextDataType = XssfDataType.ERROR;
                else if ("inlineStr".equals(cellType))
                    nextDataType = XssfDataType.INLINE_STRING;
                else if ("s".equals(cellType))
                    nextDataType = XssfDataType.SST_STRING;
                else if ("str".equals(cellType))
                    nextDataType = XssfDataType.FORMULA;
                else {
                    // Number, but almost certainly with a special style or format
                    XSSFCellStyle style = null;
                    if (stylesTable != null) {
                        if (cellStyleStr != null) {
                            int styleIndex = Integer.parseInt(cellStyleStr);
                            style = stylesTable.getStyleAt(styleIndex);
                        } else if (stylesTable.getNumCellStyles() > 0) {
                            style = stylesTable.getStyleAt(0);
                        }
                    }
                    if (style != null) {
                        this.formatIndex = style.getDataFormat();
                        this.formatString = style.getDataFormatString();
                        if (this.formatString == null)
                            this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                    }
                }
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) throws SAXException {

            if (uri != null && !uri.equals(XSSFRelation.NS_SPREADSHEETML)) {
                return;
            }

            String thisStr = null;

            // v => contents of a cell
            if (isTextTag(localName)) {
                vIsOpen = false;

                SaxCell saxCell = (SaxCell) tempRow.createCell(tempRow.getCells().size());//new SaxCell();
                tempRow.getCells().add(saxCell);

                // Process the value contents as required, now we have it all
                switch (nextDataType) {
                    case BOOLEAN:
                        char first = value.charAt(0);
                        thisStr = first == '0' ? "FALSE" : "TRUE";
                        saxCell.setCellValue(first != '0');
                        saxCell.setCellType(CellType.BOOLEAN);
                        break;

                    case ERROR:
                        thisStr = "ERROR:" + value;
                        saxCell.setCellErrorValue(value.toString().getBytes()[0]);
                        saxCell.setCellType(CellType.ERROR);
                        break;

                    case FORMULA:
                        /*if (formulasNotResults) {
                            thisStr = formula.toString();
                        } else {
                            String fv = value.toString();

                            if (this.formatString != null) {
                                try {
                                    // Try to use the value as a formattable number
                                    double d = Double.parseDouble(fv);
                                    thisStr = formatter.formatRawCellContents(d, this.formatIndex, this.formatString);
                                } catch (NumberFormatException e) {
                                    // Formula is a String result not a Numeric one
                                    thisStr = fv;
                                }
                            } else {
                                // No formatting applied, just do raw value in all cases
                                thisStr = fv;
                            }
                        }
                        saxCell.setCellFormula(thisStr);
                        saxCell.setCellType(CellType.FORMULA);*/

                        String fv = value.toString();
                        if (this.formatString != null) {
                            try {
                                // Try to use the value as a formattable number
                                double d = Double.parseDouble(fv);
                                thisStr = formatter.formatRawCellContents(d, this.formatIndex, this.formatString);
                            } catch (NumberFormatException e) {
                                // Formula is a String result not a Numeric one
                                thisStr = fv;
                            }
                        } else {
                            // No formatting applied, just do raw value in all cases
                            thisStr = fv;
                        }
                        saxCell.setCellFormula(formula.toString());
                        saxCell.setCellValue(thisStr);
                        saxCell.setCellType(CellType.FORMULA);
                        break;
                    case INLINE_STRING:
                        // TODO: Can these ever have formatting on them?
                        XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                        thisStr = rtsi.toString();
                        saxCell.setCellValue(thisStr);
                        saxCell.setCellType(CellType.STRING);
                        break;

                    case SST_STRING:
                        String sstIndex = value.toString();
                        try {
                            int idx = Integer.parseInt(sstIndex);
                            RichTextString rtss = sharedStringsTable.getItemAt(idx);
                            thisStr = rtss.toString();
                        } catch (NumberFormatException ex) {
                            logger.log(POILogger.ERROR, "Failed to parse SST index '" + sstIndex, ex);
                            thisStr = null;
                        }
                        saxCell.setCellValue(thisStr);
                        saxCell.setCellType(CellType.STRING);
                        break;

                    case NUMBER:
                        String n = value.toString();
                        if (this.formatString != null && n.length() > 0) {
                            thisStr = formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString);
                            saxCell.setCellValue(Double.parseDouble(n));
                            saxCell.setCellType(CellType.NUMERIC);
                        } else {
                            thisStr = n;
                            saxCell.setCellValue(n);
                            saxCell.setCellType(CellType.STRING);
                        }
                        break;

                    default:
                        thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                        break;
                }

                // Do we have a comment for this cell?
                checkForEmptyCellComments(EmptyCellCommentsCheckType.CELL);
                XSSFComment comment = comments != null ? comments.findCellComment(new CellAddress(cellRef)) : null;

                // Output
                //TODO output.cell(cellRef, thisStr, comment);

            } else if ("f".equals(localName)) {
                fIsOpen = false;
            } else if ("is".equals(localName)) {
                isIsOpen = false;
            } else if ("row".equals(localName)) {
                // Handle any "missing" cells which had comments attached
                checkForEmptyCellComments(EmptyCellCommentsCheckType.END_OF_ROW);

                // Finish up the row
                //TODO output.endRow(rowNum);

                // some sheets do not have rowNum set in the XML, Excel can read them so we should try to read them as well
                nextRowNum = rowNum + 1;
            } else if ("sheetData".equals(localName)) {
                // Handle any "missing" cells which had comments attached
                checkForEmptyCellComments(EmptyCellCommentsCheckType.END_OF_SHEET_DATA);

                // indicate that this sheet is now done
                //TODO output.endSheet();
                tempSheet.getRows().sort(Comparator.comparingInt(Row::getRowNum));
            } else if ("oddHeader".equals(localName) || "evenHeader".equals(localName) ||
                    "firstHeader".equals(localName)) {
                hfIsOpen = false;
                //TODO output.headerFooter(headerFooter.toString(), true, localName);
            } else if ("oddFooter".equals(localName) || "evenFooter".equals(localName) ||
                    "firstFooter".equals(localName)) {
                hfIsOpen = false;
                //TODO output.headerFooter(headerFooter.toString(), false, localName);
            }
        }

        /**
         * Captures characters only if a suitable element is open.
         * Originally was just "v"; extended for inlineStr also.
         */
        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            if (vIsOpen) {
                value.append(ch, start, length);
            }
            if (fIsOpen) {
                formula.append(ch, start, length);
            }
            if (hfIsOpen) {
                headerFooter.append(ch, start, length);
            }
        }

        /**
         * Do a check for, and output, comments in otherwise empty cells.
         */
        private void checkForEmptyCellComments(EmptyCellCommentsCheckType type) {
            if (commentCellRefs != null && !commentCellRefs.isEmpty()) {
                // If we've reached the end of the sheet data, output any
                //  comments we haven't yet already handled
                if (type == EmptyCellCommentsCheckType.END_OF_SHEET_DATA) {
                    while (!commentCellRefs.isEmpty()) {
                        outputEmptyCellComment(commentCellRefs.remove());
                    }
                    return;
                }

                // At the end of a row, handle any comments for "missing" rows before us
                if (this.cellRef == null) {
                    if (type == EmptyCellCommentsCheckType.END_OF_ROW) {
                        while (!commentCellRefs.isEmpty()) {
                            if (commentCellRefs.peek().getRow() == rowNum) {
                                outputEmptyCellComment(commentCellRefs.remove());
                            } else {
                                return;
                            }
                        }
                        return;
                    } else {
                        throw new IllegalStateException("Cell ref should be null only if there are only empty cells in the row; rowNum: " + rowNum);
                    }
                }

                CellAddress nextCommentCellRef;
                do {
                    CellAddress cellRef = new CellAddress(this.cellRef);
                    CellAddress peekCellRef = commentCellRefs.peek();
                    if (type == EmptyCellCommentsCheckType.CELL && cellRef.equals(peekCellRef)) {
                        // remove the comment cell ref from the list if we're about to handle it alongside the cell content
                        commentCellRefs.remove();
                        return;
                    } else {
                        // fill in any gaps if there are empty cells with comment mixed in with non-empty cells
                        int comparison = peekCellRef.compareTo(cellRef);
                        if (comparison > 0 && type == EmptyCellCommentsCheckType.END_OF_ROW && peekCellRef.getRow() <= rowNum) {
                            nextCommentCellRef = commentCellRefs.remove();
                            outputEmptyCellComment(nextCommentCellRef);
                        } else if (comparison < 0 && type == EmptyCellCommentsCheckType.CELL && peekCellRef.getRow() <= rowNum) {
                            nextCommentCellRef = commentCellRefs.remove();
                            outputEmptyCellComment(nextCommentCellRef);
                        } else {
                            nextCommentCellRef = null;
                        }
                    }
                } while (nextCommentCellRef != null && !commentCellRefs.isEmpty());
            }
        }

        /**
         * Output an empty-cell comment.
         */
        private void outputEmptyCellComment(CellAddress cellRef) {
            XSSFComment comment = comments.findCellComment(cellRef);
            //TODO output.cell(cellRef.formatAsString(), null, comment);
        }
    }
}
