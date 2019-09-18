package com.github.linyuzai.jpoi.excel.write.setter;

import org.apache.poi.ss.usermodel.*;

public interface ValueSetter {

    void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, Object value);
}
