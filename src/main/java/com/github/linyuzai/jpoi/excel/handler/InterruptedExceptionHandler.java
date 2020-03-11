package com.github.linyuzai.jpoi.excel.handler;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class InterruptedExceptionHandler implements ExcelExceptionHandler {

    private static final InterruptedExceptionHandler sInstance = new InterruptedExceptionHandler();

    public static InterruptedExceptionHandler getInstance() {
        return sInstance;
    }

    /**
     * @param s
     * @param r
     * @param c
     * @param cell
     * @param row
     * @param sheet
     * @param workbook
     * @param e
     * @return 是否终止
     */
    @Override
    public boolean handle(JExcelProcessor<?> context, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook, Throwable e) {
        return true;
    }
}
