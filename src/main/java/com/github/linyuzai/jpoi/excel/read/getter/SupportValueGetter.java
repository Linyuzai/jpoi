package com.github.linyuzai.jpoi.excel.read.getter;

import com.github.linyuzai.jpoi.excel.value.error.ByteError;
import com.github.linyuzai.jpoi.excel.value.formula.StringFormula;
import org.apache.poi.ss.usermodel.*;

public class SupportValueGetter extends DataFormatGetter {

    public static SupportValueGetter sInstance = new SupportValueGetter();

    public static SupportValueGetter getInstance() {
        return sInstance;
    }

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper) {
        Object cellData;
        switch (cell.getCellType()) {
            case FORMULA:
                cellData = new StringFormula(cell.getCellFormula());
                break;
            case ERROR:
                cellData = new ByteError(cell.getErrorCellValue());
                break;
            default:
                cellData = super.getValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
                break;
        }
        return cellData;
    }
}
