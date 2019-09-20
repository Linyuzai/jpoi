package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.read.JExcelAnalyzer;
import com.github.linyuzai.jpoi.excel.read.sax.SaxWorkbook;
import com.github.linyuzai.jpoi.excel.write.JExcelTransfer;
import com.github.linyuzai.jpoi.excel.write.workbook.AutoWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

public class JExcel {

    public static JExcelTransfer xls() {
        return of(new HSSFWorkbook());
    }

    public static JExcelAnalyzer xls(InputStream is) throws IOException {
        return from(new HSSFWorkbook(is));
    }

    public static JExcelTransfer xlsx() {
        return of(new XSSFWorkbook());
    }

    public static JExcelAnalyzer xlsx(InputStream is) throws IOException {
        return from(new XSSFWorkbook(is));
    }

    public static JExcelTransfer sxlsx() {
        return of(new SXSSFWorkbook());
    }

    public static JExcelAnalyzer sxlsx(InputStream is) throws IOException {
        return from(new SaxWorkbook(is));
    }

    public static JExcelTransfer auto() {
        return of(new AutoWorkbook());
    }

    public static JExcelTransfer of(Workbook workbook) {
        if (workbook == null) {
            throw new RuntimeException("workbook is null");
        }
        return new JExcelTransfer(workbook);
    }

    public static JExcelAnalyzer from(Workbook workbook) {
        if (workbook == null) {
            throw new RuntimeException("workbook is null");
        }
        return new JExcelAnalyzer(workbook);
    }
}
