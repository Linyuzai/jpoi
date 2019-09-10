package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.auto.AutoWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JExcel {

    public static JExcelTransfer xls() {
        return use(new HSSFWorkbook());
    }

    public static JExcelTransfer xlsx() {
        return use(new XSSFWorkbook());
    }

    public static JExcelTransfer sxlsx() {
        return use(new SXSSFWorkbook());
    }

    public static JExcelTransfer auto() {
        return use(new AutoWorkbook());
    }

    public static JExcelTransfer use(Workbook workbook) {
        if (workbook == null) {
            throw new RuntimeException("workbook is null");
        }
        return new JExcelTransfer(workbook);
    }
}
