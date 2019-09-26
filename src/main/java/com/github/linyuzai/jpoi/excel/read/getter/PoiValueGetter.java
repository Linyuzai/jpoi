package com.github.linyuzai.jpoi.excel.read.getter;

import org.apache.poi.ss.usermodel.*;

public class PoiValueGetter implements ValueGetter {

    private static PoiValueGetter sInstance = new PoiValueGetter();

    public static PoiValueGetter getInstance() {
        return sInstance;
    }

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        switch (cell.getCellType()) {
            case _NONE:
            case BLANK:
                return null;
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                //TODO 判断格式 DateUtil.getJavaDate();cell.getCellStyle().getDataFormat()
                return cell.getNumericCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case ERROR:
                return cell.getErrorCellValue();
        }
        return null;
    }
}
