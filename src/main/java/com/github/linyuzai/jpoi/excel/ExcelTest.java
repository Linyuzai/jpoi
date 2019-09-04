package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.adapter.SimpleDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.annotation.JExcelCell;
import com.github.linyuzai.jpoi.excel.annotation.JExcelSheet;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelTest {

    public static void main(String[] args) throws IOException {
        List<TestBean> list = new ArrayList<>();
        list.add(new TestBean("111", 1.0));
        list.add(new TestBean("222", 2.0));
        list.add(new TestBean("333", 3.0));
        list.add(new TestBean("444", 4.0));
        //JExcel.xlsx().data(list).append().data(list).write().to(new File("C:\\JExcel\\111.xlsx"));
        //JExcel.xlsx().data(list, list).write().to(new File("C:\\JExcel\\111.xlsx"));
        //SimpleDataWriteAdapter sa = new SimpleDataWriteAdapter(list, TestBean::getTestString, TestBean::getTestDouble);
        //sa.addListData(list, TestBean::getTestString);
        JExcel.auto().data(list, list).write().to(new File("C:\\JExcel\\111.xlsx"));
    }

    @JExcelSheet(name = "2222222")
    public static class TestBean {

        @JExcelCell(title = "str")
        private String testString;
        private Double testDouble;

        public TestBean(String testString, Double testDouble) {
            this.testString = testString;
            this.testDouble = testDouble;
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
    }
}
