package com.github.linyuzai.jpoi.excel.write.listener;

import com.github.linyuzai.jpoi.order.Ordered;
import org.apache.poi.ss.usermodel.*;

public interface PoiWriteListener extends Ordered {

    default void onWorkbookCreate(Workbook workbook) {

    }

    default void onSheetCreate(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {

    }

    default void onRowCreate(int r, int s, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onCellCreate(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onCellValueSet(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onRowValueSet(int r, int s, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onSheetValueSet(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {

    }

    default void onWorkbookValueSet(Workbook workbook) {

    }
}
