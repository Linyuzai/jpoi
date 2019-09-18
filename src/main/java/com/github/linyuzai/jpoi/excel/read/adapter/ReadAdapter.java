package com.github.linyuzai.jpoi.excel.read.adapter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ReadAdapter {

    Object readCell(Object value, int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook);

    Object readRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook);

    Object readSheet(List<?> rowValues, int s, Sheet sheet, Workbook workbook);

    Object readWorkbook(List<?> sheetValues, Workbook workbook);
}
