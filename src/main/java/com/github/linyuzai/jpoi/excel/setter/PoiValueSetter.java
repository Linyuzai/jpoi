package com.github.linyuzai.jpoi.excel.setter;

import org.apache.poi.ss.usermodel.*;

import java.util.Calendar;
import java.util.Date;

public class PoiValueSetter implements ValueSetter {

    private static PoiValueSetter sInstance = new PoiValueSetter();

    public static PoiValueSetter getInstance() {
        return sInstance;
    }

    @Override
    public void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, Object value) {
        if (value == null) {
            value = getNullValue(s, r, c, cell, row, sheet, drawing, workbook);
            if (value == null) {
                return;
            }
        }
        if (value.getClass() == boolean.class) {
            cell.setCellValue((boolean) value);
        } else if (value.getClass() == double.class) {
            cell.setCellValue((double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }

    String getNullValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        return null;
    }
}
