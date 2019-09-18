package com.github.linyuzai.jpoi.excel.read.adapter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public abstract class HeaderReadAdapter implements ReadAdapter {

    public Object readDataCell(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook){
        return null;
    }

    public Object readDataRow(List<Object> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook){
        return null;
    }

    public abstract int getHeaderRowCount(int sheet);

    public abstract int getHeaderCellCount(int sheet, int row);
}
