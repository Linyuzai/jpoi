package com.github.linyuzai.jpoi.excel.handler;

import org.apache.poi.ss.usermodel.*;

public interface ExcelExceptionHandler {

    boolean handle(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook, Throwable e);
}
