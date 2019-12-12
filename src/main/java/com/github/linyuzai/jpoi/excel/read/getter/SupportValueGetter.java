package com.github.linyuzai.jpoi.excel.read.getter;

import com.github.linyuzai.jpoi.excel.value.error.ByteError;
import com.github.linyuzai.jpoi.excel.value.formula.FormulaValue;
import org.apache.poi.ss.usermodel.*;

public class SupportValueGetter extends DataFormatGetter {

    public static SupportValueGetter sInstance = new SupportValueGetter();

    public static SupportValueGetter getInstance() {
        return sInstance;
    }

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper) {
        if (cell == null) {
            return null;
        }
        Object cellData;
        switch (cell.getCellType()) {
            case FORMULA:
                FormulaEvaluator formulaEvaluator = creationHelper.createFormulaEvaluator();
                CellValue cellValue = formulaEvaluator.evaluate(cell);
                cellData = new FormulaValue(cell.getCellFormula(), cellValue.getCellType() == CellType.NUMERIC ? cellValue.getNumberValue() : cellValue.getStringValue());
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
