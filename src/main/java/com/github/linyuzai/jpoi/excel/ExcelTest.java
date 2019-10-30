package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.listener.PoiListener;
import com.github.linyuzai.jpoi.excel.read.annotation.JExcelCellReader;
import com.github.linyuzai.jpoi.excel.read.annotation.JExcelSheetReader;
import com.github.linyuzai.jpoi.excel.write.adapter.WriteAdapter;
import com.github.linyuzai.jpoi.excel.write.annotation.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class ExcelTest {

    public static void main(String[] args) throws IOException {
        //JExcel.sxlsx(new FileInputStream("C:\\JExcel\\111.xlsx"));
        //read();
        //write();
        //write2();
        //read2();
        //mainWrite();
        mainRead();
    }

    private static void mainWrite() throws IOException {
        JExcel.sxlsx().data(Collections.singletonList(getExcelBean())).write().to(new File("C:\\JExcel\\Excel-Bean.xlsx"));
    }

    private static void mainRead() throws IOException {
        Object o = JExcel.xlsx(new FileInputStream("C:\\JExcel\\Excel-Bean.xlsx")).target(ExcelBean.class).read().getValue();
        System.out.println(o);
    }

    private static ExcelBean getExcelBean() {
        File file = new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg");
        byte[] buffer = null;
        try (FileInputStream fis = new FileInputStream(file);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            byte[] b = new byte[1024];
            int n;
            while ((n = fis.read(b)) != -1) {
                bos.write(b, 0, n);
            }
            fis.close();
            bos.close();
            buffer = bos.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
        }
        ExcelBean excelBean = new ExcelBean();
        excelBean._int = 101;
        excelBean._Integer = 102;
        excelBean._double = 1001.1;
        excelBean._Double = 1002.2;
        excelBean._short = 11;
        excelBean._Short = 12;
        excelBean._byte = 1;
        excelBean._Byte = 2;
        excelBean._long = 10001L;
        excelBean._Long = 10002L;
        excelBean._float = 100001f;
        excelBean._Float = 100002f;
        excelBean._char = 'a';
        excelBean._Character = 'b';
        excelBean._boolean = true;
        excelBean._Boolean = false;
        excelBean._String = "str";
        excelBean._Date = new Date();
        excelBean.comment = "Comment";
        excelBean.picture = buffer;
        excelBean.commentOfField = "Comment of field";
        excelBean.pictureOfField = buffer;
        excelBean.commentOfIndex = "Comment of index";
        String base64 = Base64.getEncoder().encodeToString(buffer);
        excelBean.pictureOfIndex = base64;
        excelBean.error = FormulaError.NA.getCode();
        excelBean.formula = "SUM(A2+B2)";
        excelBean.annotation = "annotation";
        return excelBean;
    }

    @JExcelSheetReader
    @JExcelSheetWriter(name = "Excel-Bean")
    public static class ExcelBean {
        @JExcelCellReader(title = "basic-int")
        @JExcelCellWriter(title = "basic-int")
        private int _int;
        @JExcelCellReader(title = "wrapper-integer")
        @JExcelCellWriter(title = "wrapper-integer")
        private Integer _Integer;
        @JExcelCellReader(title = "basic-double")
        @JExcelCellWriter(title = "basic-double")
        private double _double;
        @JExcelCellReader(title = "wrapper-double")
        @JExcelCellWriter(title = "wrapper-double")
        private Double _Double;
        @JExcelCellReader(title = "basic-short")
        @JExcelCellWriter(title = "basic-short")
        private short _short;
        @JExcelCellReader(title = "wrapper-short")
        @JExcelCellWriter(title = "wrapper-short")
        private Short _Short;
        @JExcelCellReader(title = "basic-byte")
        @JExcelCellWriter(title = "basic-byte")
        private byte _byte;
        @JExcelCellReader(title = "wrapper-byte")
        @JExcelCellWriter(title = "wrapper-byte")
        private Byte _Byte;
        @JExcelCellReader(title = "basic-long")
        @JExcelCellWriter(title = "basic-long")
        private long _long;
        @JExcelCellReader(title = "wrapper-long")
        @JExcelCellWriter(title = "wrapper-long")
        private Long _Long;
        @JExcelCellReader(title = "basic-float")
        @JExcelCellWriter(title = "basic-float")
        private float _float;
        @JExcelCellReader(title = "wrapper-float")
        @JExcelCellWriter(title = "wrapper-float")
        private Float _Float;
        @JExcelCellReader(title = "basic-char")
        @JExcelCellWriter(title = "basic-char")
        private char _char;
        @JExcelCellReader(title = "wrapper-character")
        @JExcelCellWriter(title = "wrapper-character")
        private Character _Character;
        @JExcelCellReader(title = "basic-boolean")
        @JExcelCellWriter(title = "basic-boolean")
        private boolean _boolean;
        @JExcelCellReader(title = "wrapper-boolean")
        @JExcelCellWriter(title = "wrapper-boolean")
        private Boolean _Boolean;
        @JExcelCellReader(title = "class-string")
        @JExcelCellWriter(title = "class-string")
        private String _String;
        @JExcelCellReader(title = "class-date")
        @JExcelCellWriter(title = "class-date", style = @JExcelCellStyle(dataFormatString = "m/d/yy"))
        private Date _Date;
        @JExcelCellReader(title = "object-comment", valueConverter = ReadCommentValueConverter.class)
        @JExcelCellWriter(title = "object-comment", valueConverter = WriteCommentValueConverter.class)
        private String comment;//注解的形式
        @JExcelCellReader(title = "object-picture", valueConverter = ReadPictureValueConverter.class)
        @JExcelCellWriter(title = "object-picture", valueConverter = WritePictureValueConverter.class)
        private byte[] picture;//图片
        @JExcelCellReader(commentOfField = "_int")
        @JExcelCellWriter(commentOfField = "_int")
        private String commentOfField;//注释，到某个字段的单元格上
        @JExcelCellReader(pictureOfFiled = "_Integer")
        @JExcelCellWriter(pictureOfFiled = "_Integer")
        private byte[] pictureOfField;//图片，到某个字段的单元格上
        @JExcelCellReader(commentOfIndex = 3)
        @JExcelCellWriter(commentOfIndex = 3)
        private String commentOfIndex;//注释，到某一列的单元格上
        @JExcelCellReader(pictureOfIndex = 4, valueConverter = ReadBase64PictureValueConverter.class)
        @JExcelCellWriter(pictureOfIndex = 4)
        private String pictureOfIndex;//base64图片，到某一列的单元格上
        @JExcelCellReader(title = "value-error")
        @JExcelCellWriter(title = "value-error", valueConverter = ErrorValueConverter.class)
        private byte error;//错误类型

        @JExcelCellWriter(title = "value-formula", valueConverter = WriteFormulaValueConverter.class)
        private String formula;//公式类型
        @JExcelCellReader(title = "value-formula", valueConverter = ReadFormulaValueConverter.class)
        private Double formulaValue;//公式类型

        private String annotation;//测试annotationOnly属性

        public int get_int() {
            return _int;
        }

        public void set_int(int _int) {
            this._int = _int;
        }

        public Integer get_Integer() {
            return _Integer;
        }

        public void set_Integer(Integer _Integer) {
            this._Integer = _Integer;
        }

        public double get_double() {
            return _double;
        }

        public void set_double(double _double) {
            this._double = _double;
        }

        public Double get_Double() {
            return _Double;
        }

        public void set_Double(Double _Double) {
            this._Double = _Double;
        }

        public short get_short() {
            return _short;
        }

        public void set_short(short _short) {
            this._short = _short;
        }

        public Short get_Short() {
            return _Short;
        }

        public void set_Short(Short _Short) {
            this._Short = _Short;
        }

        public byte get_byte() {
            return _byte;
        }

        public void set_byte(byte _byte) {
            this._byte = _byte;
        }

        public Byte get_Byte() {
            return _Byte;
        }

        public void set_Byte(Byte _Byte) {
            this._Byte = _Byte;
        }

        public long get_long() {
            return _long;
        }

        public void set_long(long _long) {
            this._long = _long;
        }

        public Long get_Long() {
            return _Long;
        }

        public void set_Long(Long _Long) {
            this._Long = _Long;
        }

        public float get_float() {
            return _float;
        }

        public void set_float(float _float) {
            this._float = _float;
        }

        public Float get_Float() {
            return _Float;
        }

        public void set_Float(Float _Float) {
            this._Float = _Float;
        }

        public char get_char() {
            return _char;
        }

        public void set_char(char _char) {
            this._char = _char;
        }

        public Character get_Character() {
            return _Character;
        }

        public void set_Character(Character _Character) {
            this._Character = _Character;
        }

        public boolean is_boolean() {
            return _boolean;
        }

        public void set_boolean(boolean _boolean) {
            this._boolean = _boolean;
        }

        public Boolean get_Boolean() {
            return _Boolean;
        }

        public void set_Boolean(Boolean _Boolean) {
            this._Boolean = _Boolean;
        }

        public String get_String() {
            return _String;
        }

        public void set_String(String _String) {
            this._String = _String;
        }

        public Date get_Date() {
            return _Date;
        }

        public void set_Date(Date _Date) {
            this._Date = _Date;
        }

        public byte[] getPicture() {
            return picture;
        }

        public void setPicture(byte[] picture) {
            this.picture = picture;
        }

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }

        public String getCommentOfField() {
            return commentOfField;
        }

        public void setCommentOfField(String commentOfField) {
            this.commentOfField = commentOfField;
        }

        public byte[] getPictureOfField() {
            return pictureOfField;
        }

        public void setPictureOfField(byte[] pictureOfField) {
            this.pictureOfField = pictureOfField;
        }

        public String getCommentOfIndex() {
            return commentOfIndex;
        }

        public void setCommentOfIndex(String commentOfIndex) {
            this.commentOfIndex = commentOfIndex;
        }

        public String getPictureOfIndex() {
            return pictureOfIndex;
        }

        public void setPictureOfIndex(String pictureOfIndex) {
            this.pictureOfIndex = pictureOfIndex;
        }

        public byte getError() {
            return error;
        }

        public void setError(byte error) {
            this.error = error;
        }

        public String getFormula() {
            return formula;
        }

        public void setFormula(String formula) {
            this.formula = formula;
        }

        public Double getFormulaValue() {
            return formulaValue;
        }

        public void setFormulaValue(Double formulaValue) {
            this.formulaValue = formulaValue;
        }

        public String getAnnotation() {
            return annotation;
        }

        public void setAnnotation(String annotation) {
            this.annotation = annotation;
        }

        @JExcelCellWriter(commentOfField = "formula")
        public String testMethod() {
            return get_String();
        }
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
        JExcel.sxlsx().setWriteAdapter(new WriteAdapter() {
            @Override
            public Object getData(int sheet, int row, int cell) {
                if (row == 0) {
                    if (cell == 0) {
                        return "Test String";
                    } else if (cell == 1) {
                        return "Test Double";
                    } else if (cell == 2) {
                        return "Test Date";
                    }
                } else {
                    if (cell == 0) {
                        return list.get(row).getTestString();
                    } else if (cell == 1) {
                        return list.get(row).getTestDouble();
                    } else if (cell == 2) {
                        return list.get(row).getTestDate();
                    }
                }
                return null;
            }

            @Override
            public String getSheetName(int sheet) {
                return "Sheet1";
            }

            @Override
            public int getSheetCount() {
                return 1;
            }

            @Override
            public int getRowCount(int sheet) {
                return list.size() + 1;
            }

            @Override
            public int getCellCount(int sheet, int row) {
                return 3;
            }
        }).write().to(new File("C:\\JExcel\\test-bean2.xlsx"));
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

        //@JExcelCellWriter(autoSize = false, valueConverter = WritePictureValueConverter.class)
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
