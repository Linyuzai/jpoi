package com.github.linyuzai.jpoi.excel.handler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class StdoutExceptionHandler implements ExcelExceptionHandler {

    private static final StdoutExceptionHandler sInstance = new StdoutExceptionHandler();

    public static StdoutExceptionHandler getInstance() {
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
    public boolean handle(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook, Throwable e) {
        e.printStackTrace();
        return false;
    }
}
