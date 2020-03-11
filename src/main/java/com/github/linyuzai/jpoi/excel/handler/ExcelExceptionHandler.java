package com.github.linyuzai.jpoi.excel.handler;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import org.apache.poi.ss.usermodel.*;

public interface ExcelExceptionHandler {

    boolean handle(JExcelProcessor<?> context, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook, Throwable e);
}
