package com.github.linyuzai.jpoi.excel.read.getter;

import com.github.linyuzai.jpoi.excel.value.format.NumericDataFormatValue;
import org.apache.poi.ss.usermodel.*;

public class DataFormatGetter extends PoiValueGetter {

    private static DataFormatter dataFormatter = new DataFormatter();
    private static DataFormatGetter sInstance = new DataFormatGetter();

    public static DataFormatGetter getInstance() {
        return sInstance;
    }

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper) {
        if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else {
                short dataFormat = cell.getCellStyle().getDataFormat();
                if (dataFormat == 0) {
                    return cell.getNumericCellValue();
                } else {
                    String formatValue = getDataFormatter(s, r, c, cell, row, sheet, drawing, workbook, creationHelper).formatCellValue(cell);
                    return new NumericDataFormatValue(cell.getNumericCellValue(), formatValue, dataFormat);
                }
            }
        }
        return super.getValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
    }

    public DataFormatter getDataFormatter(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper) {
        return dataFormatter;
    }
}
