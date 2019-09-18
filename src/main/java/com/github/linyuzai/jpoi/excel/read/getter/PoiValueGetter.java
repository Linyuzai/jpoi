package com.github.linyuzai.jpoi.excel.read.getter;

import org.apache.poi.ss.usermodel.*;

public class PoiValueGetter implements ValueGetter {

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        switch (cell.getCellType()) {
            case _NONE:
            case BLANK:
                return null;
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                //TODO 判断日期 DateUtil.getJavaDate();
                return cell.getNumericCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case ERROR:
                return cell.getErrorCellValue();
        }
        return null;
    }
}
