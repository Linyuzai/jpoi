package com.github.linyuzai.jpoi.excel.read.adapter;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;

@SuppressWarnings("unchecked")
public class DirectListReadAdapter extends ListReadAdapter {

    @Override
    public Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return new HashMap<Integer, Object>();
    }

    @Override
    public void adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        ((Map<Integer, Object>) cellContainer).put(c, value);
    }

    @Override
    public Object getValue() {
        return getReadMap().values().stream()
                .collect(Collectors.collectingAndThen(Collectors.toList(), r -> r.stream()
                        .map(rm -> rm.values().stream()
                                .collect(Collectors.collectingAndThen(Collectors.toList(), c -> c.stream()
                                        .map(cm -> new ArrayList<>(((Map<Integer, Object>) cm).values()))
                                        .collect(Collectors.toList()))))))
                .collect(Collectors.toList());
    }
}
