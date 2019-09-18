package com.github.linyuzai.jpoi.excel.read.adapter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public abstract class HeaderReadAdapter implements ReadAdapter {

    @Override
    public Object readCell(Object value, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        if (c < getHeaderCellCount(s, r)) {
            return readHeaderCell(s, r, c, cell, row, sheet, workbook);
        } else {
            return readDataCell(s, r, c, cell, row, sheet, workbook);
        }
    }

    @Override
    public Object readRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook) {
        if (r < getHeaderRowCount(s)) {
            return readHeaderRow(cellValues, s, r, row, sheet, workbook);
        } else {
            return readDataRow(cellValues, s, r, row, sheet, workbook);
        }
    }

    public abstract Object readHeaderCell(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook);

    public abstract Object readHeaderRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook);

    public abstract Object readDataCell(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook);

    public abstract Object readDataRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook);

    public abstract int getHeaderRowCount(int sheet);

    public abstract int getHeaderCellCount(int sheet, int row);
}
