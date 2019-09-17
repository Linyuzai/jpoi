package com.github.linyuzai.jpoi.excel.adapter.read;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ReadAdapter {

    Object readCell(int s, int r, int c, Cell cell, Row row, Sheet sheet, Workbook workbook);

    Object readRow(List<Object> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook);

    Object readSheet(List<Object> rowValues, int s, Sheet sheet, Workbook workbook);

    Object readWorkbook(List<Object> sheetValues, Workbook workbook);
}
