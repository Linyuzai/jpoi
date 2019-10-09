package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.converter.ReadPictureValueConverter;
import com.github.linyuzai.jpoi.excel.listener.PoiListener;
import com.github.linyuzai.jpoi.excel.read.annotation.JExcelCellReader;
import com.github.linyuzai.jpoi.excel.read.annotation.JExcelSheetReader;
import com.github.linyuzai.jpoi.excel.write.annotation.*;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

public class ExcelTest {

    public static void main(String[] args) throws IOException {
        //JExcel.sxlsx(new FileInputStream("C:\\JExcel\\111.xlsx"));
        //read();
        //write();
        write2();
        //read2();
    }

    public static void read() throws IOException {
        Object o = JExcel.xlsx(new FileInputStream("C:\\JExcel\\111.xlsx")).target(TestBean.class).read().getValue();
        System.out.println(o);
    }

    public static void read2() throws IOException {
        Object o = JExcel.xlsx(new FileInputStream("C:\\JExcel\\test-bean2.xlsx")).target(TestBean2.class).addPoiReadListener(new PoiListener() {
            @Override
            public void onRowStart(int r, int s, Row row, Sheet sheet, Workbook workbook) {
                System.out.println("onRowStart:getHeight-->" + row.getHeight());
                System.out.println("onRowStart:getHeightInPoints-->" + row.getHeightInPoints());
                System.out.println("onRowStart:getZeroHeight-->" + row.getZeroHeight());
                //printStyle(row.getRowStyle(), workbook);
            }

            @Override
            public void onCellStart(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {
                printStyle(cell.getCellStyle(), workbook);
            }
        }).read().getValue();
        System.out.println(o);
    }

    private static void printStyle(CellStyle style, Workbook workbook) {
        System.out.println("getAlignment-->" + style.getAlignment());
        System.out.println("getBorderBottom-->" + style.getBorderBottom());
        System.out.println("getBottomBorderColor-->" + style.getBottomBorderColor());
        System.out.println("getBorderLeft-->" + style.getBorderLeft());
        System.out.println("getLeftBorderColor-->" + style.getLeftBorderColor());
        System.out.println("getBorderRight-->" + style.getBorderRight());
        System.out.println("getRightBorderColor-->" + style.getRightBorderColor());
        System.out.println("getBorderTop-->" + style.getBorderTop());
        System.out.println("getTopBorderColor-->" + style.getTopBorderColor());
        System.out.println("getDataFormat-->" + style.getDataFormat());
        System.out.println("getDataFormatString-->" + style.getDataFormatString());
        System.out.println("getFillPattern-->" + style.getFillPattern());
        System.out.println("getFillBackgroundColor-->" + style.getFillBackgroundColor());
        System.out.println("getFillForegroundColor-->" + style.getFillForegroundColor());
        System.out.println("getHidden-->" + style.getHidden());
        System.out.println("getIndention-->" + style.getIndention());
        System.out.println("getLocked-->" + style.getLocked());
        System.out.println("getQuotePrefixed-->" + style.getQuotePrefixed());
        System.out.println("getRotation-->" + style.getRotation());
        System.out.println("getShrinkToFit-->" + style.getShrinkToFit());
        System.out.println("getWrapText-->" + style.getWrapText());
        System.out.println("getVerticalAlignment-->" + style.getVerticalAlignment());
        Font font = workbook.getFontAt(style.getFontIndexAsInt());
        System.out.println("font:getFontName-->" + font.getFontName());
        System.out.println("font:getBold-->" + font.getBold());
        System.out.println("font:getCharSet-->" + font.getCharSet());
        System.out.println("font:getColor-->" + font.getColor());
        System.out.println("font:getFontHeight-->" + font.getFontHeight());
        System.out.println("font:getFontHeightInPoints-->" + font.getFontHeightInPoints());
        System.out.println("font:getItalic-->" + font.getItalic());
        System.out.println("font:getStrikeout-->" + font.getStrikeout());
        System.out.println("font:getTypeOffset-->" + font.getTypeOffset());
        System.out.println("font:getUnderline-->" + font.getUnderline());
    }

    public static void write2() throws IOException {
        List<TestBean2> list = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            list.add(new TestBean2(String.valueOf(i), i * 1.0, new Date()));
        }
        JExcel.sxlsx().data(list).write().to(new File("C:\\JExcel\\test-bean2.xlsx"));
    }

    public static void write() throws IOException {
        List<TestBean> list = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            list.add(new TestBean(UUID.randomUUID().toString(), 1.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg"), "comment"));
        }
        //list.add(new TestBean("11111111111111111111111111111111111111", 1.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //list.add(new TestBean("222", 2.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //list.add(new TestBean("333", 3.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //list.add(new TestBean("444", 4.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //JExcel.xlsx().data(list).append().data(list).write().to(new File("C:\\JExcel\\111.xlsx"));
        //JExcel.xlsx().data(list, list).write().to(new File("C:\\JExcel\\111.xlsx"));
        //SimpleDataWriteAdapter sa = new SimpleDataWriteAdapter(list, TestBean::getTestString, TestBean::getTestDouble);
        //sa.addListData(list, TestBean::getTestString);
        /*CellStyle style;
        style.setAlignment();
        style.setBorderBottom();
        style.setBottomBorderColor();
        style.setBorderLeft();
        style.setLeftBorderColor();
        style.setBorderRight();
        style.setRightBorderColor();
        style.setBorderTop();
        style.setTopBorderColor();
        style.setDataFormat();
        style.setFillPattern();
        style.setFillBackgroundColor();
        style.setFillForegroundColor();
        style.setFont();
        style.setHidden();
        style.setIndention();
        style.setLocked();
        style.setQuotePrefixed();
        style.setRotation();
        style.setShrinkToFit();
        style.setWrapText();
        style.setVerticalAlignment();*/

        /*Font font;
        font.setBold();
        font.setCharSet();
        font.setCharSet();
        font.setColor();
        font.setFontHeight();
        font.setFontHeightInPoints();
        font.setFontName();
        font.setItalic();
        font.setStrikeout();
        font.setTypeOffset();
        font.setUnderline();*/
        /*Row row;
        row.setHeight();
        row.setHeightInPoints();
        row.setZeroHeight();
        row.setRowStyle();*/
        JExcel.sxlsx()//.async()
                /*.writeAdapter(new SimpleDataWriteAdapter(list) {
                    @Override
                    public int getHeaderRowCount(int sheet) {
                        return 2;
                    }

                    @Override
                    public Object getHeaderRowContent(int sheet, int row, int cell, int realRow, int realCell) {
                        if (realRow == 0) {
                            return "test";
                        }
                        return super.getHeaderRowContent(sheet, row, cell, realRow, realCell);
                    }

                    @Override
                    public void onSheetCreate(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, getCellCount(s, 0) - 1));
                        super.onSheetCreate(s, sheet, drawing, workbook);
                    }
                })*/
                .data(list)
                .write().to(new File("C:\\JExcel\\111.xlsx"));
        System.out.println("-------------------------------------------");
    }

    @JExcelSheetReader
    @JExcelSheetWriter(name = "test-bean2", style = @JExcelRowStyle(height = 1200,
            cellStyle = @JExcelCellStyle(horizontalAlignment = HorizontalAlignment.RIGHT,
                    border = @JExcelCellBorder(top = BorderStyle.DOUBLE),
                    font = @JExcelCellFont(color = Font.COLOR_RED))))
    public static class TestBean2 {

        @JExcelCellReader(title = "Test String")
        @JExcelCellWriter(title = "Test String", style = @JExcelCellStyle(horizontalAlignment = HorizontalAlignment.LEFT))
        private String testString;

        @JExcelCellReader(title = "Test Double")
        @JExcelCellWriter(title = "Test Double", style = @JExcelCellStyle(border = @JExcelCellBorder(top = BorderStyle.DOUBLE)))
        private Double testDouble;

        @JExcelCellReader(title = "Test Date")
        @JExcelCellWriter(title = "Test Date")
        private Date testDate;

        public TestBean2() {
        }

        public TestBean2(String testString, Double testDouble, Date testDate) {
            this.testString = testString;
            this.testDouble = testDouble;
            this.testDate = testDate;
        }

        public String getTestString() {
            return testString;
        }

        public void setTestString(String testString) {
            this.testString = testString;
        }

        public Double getTestDouble() {
            return testDouble;
        }

        public void setTestDouble(Double testDouble) {
            this.testDouble = testDouble;
        }

        public Date getTestDate() {
            return testDate;
        }

        public void setTestDate(Date testDate) {
            this.testDate = testDate;
        }
    }

    @JExcelSheetReader(annotationOnly = true)
    @JExcelSheetWriter(name = "2222222", annotationOnly = true, style = @JExcelRowStyle(
            height = 1,
            heightInPoints = 1f,
            zeroHeight = true,
            cellStyle = @JExcelCellStyle
    ))
    public static class TestBean {

        //@JExcelCellReader(title = "string11111")
        @JExcelCellWriter(title = "string11111", style = @JExcelCellStyle(
                horizontalAlignment = HorizontalAlignment.CENTER,
                verticalAlignment = VerticalAlignment.BOTTOM,
                dataFormat = 0,
                hidden = true,
                indention = 0,
                locked = true,
                quotePrefixed = true,
                rotation = 1,
                shrinkToFit = true,
                wrapText = true,
                border = @JExcelCellBorder(
                        top = BorderStyle.DASH_DOT,
                        topColor = 1,
                        right = BorderStyle.DASH_DOT,
                        rightColor = 1,
                        bottom = BorderStyle.DASH_DOT,
                        bottomColor = 1,
                        left = BorderStyle.DASH_DOT,
                        leftColor = 1),
                fill = @JExcelCellFill(
                        pattern = FillPatternType.ALT_BARS,
                        foregroundColor = 1,
                        backgroundColor = 1),
                font = @JExcelCellFont(
                        bold = true,
                        charSet = 1,
                        color = 1,
                        fontHeight = 1,
                        fontHeightInPoints = 1,
                        fontName = "1",
                        italic = true,
                        strikeout = true,
                        typeOffset = 1,
                        underline = 1)
        ))
        private String testString;

        private Double testDouble;

        //@JExcelCellWriter(autoSize = false, valueConverter = PictureValueConverter.class)
        @JExcelCellWriter(pictureOfFiled = "testString")
        private File file;

        //@JExcelCellReader(pictureOfFiled = "testString")
        @JExcelCellReader(title = "string11111", valueConverter = ReadPictureValueConverter.class)
        private byte[] bytes;

        @JExcelCellReader(commentOfField = "bytes")
        @JExcelCellWriter(commentOfField = "testString")
        private String comment;

        public TestBean() {
        }

        public TestBean(String testString, Double testDouble, File file, String comment) {
            this.testString = testString;
            this.testDouble = testDouble;
            this.file = file;
            this.comment = comment;
        }

        public String getTestString() {
            return testString;
        }

        public void setTestString(String testString) {
            this.testString = testString;
        }

        public Double getTestDouble() {
            return testDouble;
        }

        public void setTestDouble(Double testDouble) {
            this.testDouble = testDouble;
        }

        public File getFile() {
            return file;
        }

        public void setFile(File file) {
            this.file = file;
        }

        public byte[] getBytes() {
            return bytes;
        }

        public void setBytes(byte[] bytes) {
            this.bytes = bytes;
        }

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }
    }
}
