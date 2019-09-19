package com.github.linyuzai.jpoi.excel.read.adapter;

import java.util.*;

public abstract class ListReadAdapter<T> extends HeaderReadAdapter {

    private Map<Integer, Map<Integer, T>> readMap;

    public ListReadAdapter() {
        readMap = createMap(-1);
    }

    public Map<Integer, Map<Integer, T>> getReadMap() {
        return readMap;
    }

    @Override
    public void readOverlappingCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }

    @Override
    public void readCellHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }

    @Override
    public void readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }

    @Override
    public void readDataCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        Map<Integer, T> rowMap = readMap.computeIfAbsent(s, k -> createMap(s));
        T cellContainer = rowMap.get(r);
        if (cellContainer == null) {
            cellContainer = createContainer(value, s, r, c, sCount, rCount, cCount);
            rowMap.put(r, cellContainer);
        }
        adaptValue(cellContainer, value, s, r, c, sCount, rCount, cCount);
    }

    public abstract T createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract void adaptValue(T cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    @Override
    public int getHeaderRowCount(int sheet) {
        return 0;
    }

    @Override
    public int getHeaderCellCount(int sheet, int row) {
        return 0;
    }

    @Override
    public Object getValue() {
        return null;
    }

    public <M> Map<Integer, M> createMap(int s) {
        return new HashMap<>();
    }
}
