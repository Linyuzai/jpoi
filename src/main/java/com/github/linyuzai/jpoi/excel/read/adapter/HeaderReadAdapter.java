package com.github.linyuzai.jpoi.excel.read.adapter;

public abstract class HeaderReadAdapter implements ReadAdapter {

    @Override
    public Object readCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable {
        int headerCellCount = getHeaderCellCount(s, r);
        int headerRowCount = getHeaderRowCount(s);
        if (c < headerCellCount && r < headerRowCount) {
            return readOverlappingCell(value, s, r, c, sCount, rCount, cCount);
        } else if (c < headerCellCount) {
            return readCellHeaderCell(value, s, r, c, sCount, rCount, cCount);
        } else if (r < headerRowCount) {
            return readRowHeaderCell(value, s, r, c, sCount, rCount, cCount);
        } else {
            return readDataCell(value, s, r, c, sCount, rCount, cCount);
        }
    }

    public abstract Object readOverlappingCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract Object readCellHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract Object readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract Object readDataCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable;

    public abstract int getHeaderRowCount(int sheet);

    public abstract int getHeaderCellCount(int sheet, int row);
}
