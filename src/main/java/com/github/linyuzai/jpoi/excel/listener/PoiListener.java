package com.github.linyuzai.jpoi.excel.listener;

import com.github.linyuzai.jpoi.order.Ordered;
import org.apache.poi.ss.usermodel.*;

public interface PoiListener extends Ordered {

    void onWorkbookCreate(Workbook workbook);

    void onSheetCreate(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook);

    void onRowCreate(int r, int s, Row row, Sheet sheet, Workbook workbook);

    void onCellCreate(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook);
}
