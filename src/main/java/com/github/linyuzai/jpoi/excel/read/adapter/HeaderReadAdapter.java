package com.github.linyuzai.jpoi.excel.read.adapter;

public abstract class HeaderReadAdapter implements ReadAdapter {

    @Override
    public void readCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        int headerCellCount = getHeaderCellCount(s, r);
        int headerRowCount = getHeaderRowCount(s);
        if (c < headerCellCount && r < headerRowCount) {
            readOverlappingCell(value, s, r, c, sCount, rCount, cCount);
        } else if (c < headerCellCount) {
            readCellHeaderCell(value, s, r, c, sCount, rCount, cCount);
        } else if (r < headerRowCount) {
            readRowHeaderCell(value, s, r, c, sCount, rCount, cCount);
        } else {
            readDataCell(value, s, r, c, sCount, rCount, cCount);
        }
    }

    public abstract void readOverlappingCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract void readCellHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract void readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract void readDataCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract int getHeaderRowCount(int sheet);

    public abstract int getHeaderCellCount(int sheet, int row);
}
