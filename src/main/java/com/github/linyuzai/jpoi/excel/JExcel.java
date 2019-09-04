package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.auto.AutoWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JExcel {

    public static JExcelTransfer xls() {
        return new JExcelTransfer(new HSSFWorkbook());
    }

    public static JExcelTransfer xlsx() {
        return new JExcelTransfer(new XSSFWorkbook());
    }

    public static JExcelTransfer sxlsx() {
        return new JExcelTransfer(new SXSSFWorkbook());
    }

    public static JExcelTransfer auto() {
        return new JExcelTransfer(new AutoWorkbook());
    }

    public static JExcelTransfer of(Workbook workbook) {
        if (workbook == null) {
            throw new RuntimeException("workbook is null");
        }
        return new JExcelTransfer(workbook);
    }
}
