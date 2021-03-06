package com.github.linyuzai.jpoi.excel.write.adapter;

public interface WriteAdapter {

    Object getData(int sheet, int row, int cell);

    String getSheetName(int sheet);

    int getSheetCount();

    int getRowCount(int sheet);

    int getCellCount(int sheet, int row);
}
