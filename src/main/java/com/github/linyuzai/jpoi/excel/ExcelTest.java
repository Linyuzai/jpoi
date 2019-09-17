package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.adapter.write.SimpleDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.annotation.JExcelCell;
import com.github.linyuzai.jpoi.excel.annotation.JExcelSheet;
import com.github.linyuzai.jpoi.excel.converter.PictureValueConverter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class ExcelTest {

    public static void main(String[] args) throws IOException {
        List<TestBean> list = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            list.add(new TestBean(UUID.randomUUID().toString(), 1.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        }
        //list.add(new TestBean("11111111111111111111111111111111111111", 1.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //list.add(new TestBean("222", 2.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //list.add(new TestBean("333", 3.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //list.add(new TestBean("444", 4.0, new File("C:\\Users\\tangh\\Desktop\\image-nova2.jpg")));
        //JExcel.xlsx().data(list).append().data(list).write().to(new File("C:\\JExcel\\111.xlsx"));
        //JExcel.xlsx().data(list, list).write().to(new File("C:\\JExcel\\111.xlsx"));
        //SimpleDataWriteAdapter sa = new SimpleDataWriteAdapter(list, TestBean::getTestString, TestBean::getTestDouble);
        //sa.addListData(list, TestBean::getTestString);
        JExcel.sxlsx()
                .writeAdapter(new SimpleDataWriteAdapter(list) {
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
                })
                //.data(list)
                .write().to(new File("C:\\JExcel\\111.xlsx"));
    }

    @JExcelSheet(name = "2222222")
    public static class TestBean {

        @JExcelCell(title = "string11111")
        private String testString;
        private Double testDouble;
        @JExcelCell(autoSize = false, valueConverter = PictureValueConverter.class)
        private File file;

        public TestBean(String testString, Double testDouble, File file) {
            this.testString = testString;
            this.testDouble = testDouble;
            this.file = file;
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
    }
}
