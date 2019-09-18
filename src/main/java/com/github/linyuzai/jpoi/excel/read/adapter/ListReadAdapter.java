package com.github.linyuzai.jpoi.excel.read.adapter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class ListReadAdapter extends HeaderReadAdapter {

    @Override
    public Object readHeaderCell(Object value, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        return value;
    }

    @Override
    public Object readHeaderRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook) {
        return cellValues;
    }

    @Override
    public Object readDataCell(Object value, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        return value;
    }

    @Override
    public Object readDataRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook) {
        return cellValues;
    }

    @Override
    public int getHeaderRowCount(int sheet) {
        return 0;
    }

    @Override
    public int getHeaderCellCount(int sheet, int row) {
        return 0;
    }

    @Override
    public Object readSheet(List<?> rowValues, int s, Sheet sheet, Workbook workbook) {
        return rowValues;
    }

    @Override
    public Object readWorkbook(List<?> sheetValues, Workbook workbook) {
        return sheetValues;
    }
}
