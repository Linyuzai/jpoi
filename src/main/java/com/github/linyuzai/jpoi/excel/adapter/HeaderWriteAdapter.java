package com.github.linyuzai.jpoi.excel.adapter;

public abstract class HeaderWriteAdapter implements WriteAdapter {

    @Override
    public Object getData(int sheet, int row, int cell) {
        int r = row - getHeaderRowCount(sheet);
        int c = cell - getHeaderCellCount(sheet, row);
        if (row < getHeaderRowCount(sheet)) {
            return getHeaderRowContent(sheet, r, c, row, cell);
        }
        if (cell < getHeaderCellCount(sheet, row)) {
            return getHeaderCellContent(sheet, r, c, row, cell);
        }
        return getDataCalculateHeader(sheet, r, c, row, cell);
    }

    @Override
    public int getRowCount(int sheet) {
        return getHeaderRowCount(sheet) + getDataRowCount(sheet);
    }

    @Override
    public int getCellCount(int sheet, int row) {
        return getHeaderCellCount(sheet, row) + getDataCellCount(sheet, row);
    }

    public abstract int getHeaderRowCount(int sheet);

    public abstract int getHeaderCellCount(int sheet, int row);

    public abstract int getDataRowCount(int sheet);

    public abstract int getDataCellCount(int sheet, int row);

    public abstract Object getHeaderRowContent(int sheet, int row, int cell, int realRow, int realCell);

    public abstract Object getHeaderCellContent(int sheet, int row, int cell, int realRow, int realCell);

    public abstract Object getDataCalculateHeader(int sheet, int row, int cell, int realRow, int realCell);
}
