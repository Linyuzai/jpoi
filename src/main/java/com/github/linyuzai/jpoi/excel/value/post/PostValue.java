package com.github.linyuzai.jpoi.excel.value.post;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import org.apache.poi.ss.usermodel.*;

public interface PostValue {

    int getSheetIndex();

    void setSheetIndex(int sheetIndex);

    int getRowIndex();

    void setRowIndex(int rowIndex);

    int getCellIndex();

    void setCellIndex(int cellIndex);

    Sheet getSheet();

    void setSheet(Sheet sheet);

    Row getRow();

    void setRow(Row row);

    Cell getCell();

    void setCell(Cell cell);

    Drawing<?> getDrawing();

    void setDrawing(Drawing<?> drawing);

    CreationHelper getCreationHelper();

    void setCreationHelper(CreationHelper creationHelper);

    Object getValue();

    ValueConverter getValueConverter();
}
