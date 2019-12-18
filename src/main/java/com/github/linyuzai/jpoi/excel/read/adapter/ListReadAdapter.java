package com.github.linyuzai.jpoi.excel.read.adapter;

import java.util.*;
import java.util.stream.Collectors;

public abstract class ListReadAdapter extends HeaderReadAdapter {

    private Map<Integer, Map<Integer, Object>> readMap;

    public ListReadAdapter() {
        readMap = createMap(-1);
    }

    public Map<Integer, Map<Integer, Object>> getReadMap() {
        return readMap;
    }

    @Override
    public void readOverlappingCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        readRowHeaderCell(value, s, r, c, sCount, rCount, cCount);
    }

    @Override
    public void readCellHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }

    @Override
    public void readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }

    @Override
    public void readDataCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable {
        Map<Integer, Object> rowMap = readMap.computeIfAbsent(s, k -> createMap(s));
        Object cellContainer = rowMap.get(r);
        if (cellContainer == null) {
            cellContainer = createContainer(value, s, r, c, sCount, rCount, cCount);
            rowMap.put(r, cellContainer);
        }
        adaptValue(cellContainer, value, s, r, c, sCount, rCount, cCount);
    }

    public abstract Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract void adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable;

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
        return getReadMap().values().stream()
                .collect(Collectors.collectingAndThen(Collectors.toList(), r -> r.stream()
                        .map(rm -> rm.values().stream()
                                .collect(Collectors.collectingAndThen(Collectors.toList(), ArrayList::new)))))
                .collect(Collectors.toList());
    }

    public <M> Map<Integer, M> createMap(int s) {
        return new HashMap<>();
    }
}
