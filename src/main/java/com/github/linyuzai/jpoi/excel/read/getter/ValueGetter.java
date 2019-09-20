package com.github.linyuzai.jpoi.excel.read.getter;

import org.apache.poi.ss.usermodel.*;

public interface ValueGetter {

    Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook);
}
