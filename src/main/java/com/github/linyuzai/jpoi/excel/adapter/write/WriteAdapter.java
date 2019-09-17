package com.github.linyuzai.jpoi.excel.adapter.write;

public interface WriteAdapter {

    Object getData(int sheet, int row, int cell);

    String getSheetName(int sheet);

    int getSheetCount();

    int getRowCount(int sheet);

    int getCellCount(int sheet, int row);
}
