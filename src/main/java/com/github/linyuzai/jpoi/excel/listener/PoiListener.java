package com.github.linyuzai.jpoi.excel.listener;

import com.github.linyuzai.jpoi.support.SupportOrder;
import org.apache.poi.ss.usermodel.*;

/**
 * start is create for write or get for read(start of handle)
 * <p>
 * end is value set for write or get for read(end of handle)
 */
public interface PoiListener extends SupportOrder {

    default void onWorkbookStart(Workbook workbook) {

    }

    default void onSheetStart(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {

    }

    default void onRowStart(int r, int s, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onCellStart(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onCellEnd(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onRowEnd(int r, int s, Row row, Sheet sheet, Workbook workbook) {

    }

    default void onSheetEnd(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {

    }

    default void onWorkbookEnd(Workbook workbook) {

    }
}
