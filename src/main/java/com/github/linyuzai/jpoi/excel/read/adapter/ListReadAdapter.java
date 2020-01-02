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
    public Object readOverlappingCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return readRowHeaderCell(value, s, r, c, sCount, rCount, cCount);
    }

    @Override
    public Object readCellHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return null;
    }

    @Override
    public Object readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return null;
    }

    @Override
    public Object readDataCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable {
        Map<Integer, Object> rowMap = readMap.computeIfAbsent(s, k -> createMap(s));
        Object cellContainer = rowMap.get(r);
        if (cellContainer == null) {
            cellContainer = createContainer(value, s, r, c, sCount, rCount, cCount);
            if (cellContainer == null) {
                return value;
            }
            rowMap.put(r, cellContainer);
        }
        return adaptValue(cellContainer, value, s, r, c, sCount, rCount, cCount);
    }

    public abstract Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount);

    public abstract Object adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable;

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
