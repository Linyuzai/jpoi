package com.github.linyuzai.jpoi.excel.listener;

import com.github.linyuzai.jpoi.order.Ordered;
import org.apache.poi.ss.usermodel.*;

/**
 * start is create for write,get for read
 * <p>
 * end is value set for write,get for read
 */
public interface PoiListener extends Ordered {

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
