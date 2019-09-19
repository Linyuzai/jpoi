package com.github.linyuzai.jpoi.excel.read.adapter;

import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;

public class DirectListReadAdapter extends ListReadAdapter<Map<Integer, Object>> {

    @Override
    public Map<Integer, Object> createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return new HashMap<>();
    }

    @Override
    public void adaptValue(Map<Integer, Object> cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        cellContainer.put(c, value);
    }

    @Override
    public Object getValue() {
        return getReadMap().values().stream()
                .collect(Collectors.collectingAndThen(Collectors.toList(), r -> r.stream()
                        .map(rm -> rm.values().stream()
                                .collect(Collectors.collectingAndThen(Collectors.toList(), c -> c.stream()
                                        .map(Map::values).collect(Collectors.toList()))))))
                .collect(Collectors.toList());
    }
}
