package com.github.linyuzai.jpoi.excel.handler;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
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
     * @return 是否终止
     */
    @Override
    public boolean handle(JExcelProcessor<?> context, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook, Throwable e) {
        e.printStackTrace();
        return false;
    }
}
