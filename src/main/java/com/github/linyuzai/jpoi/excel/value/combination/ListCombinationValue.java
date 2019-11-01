package com.github.linyuzai.jpoi.excel.value.combination;

import java.util.ArrayList;
import java.util.List;

public class ListCombinationValue implements CombinationValue {

    private List<Object> values = new ArrayList<>();

    @Override
    public Object getValue(Object key) {
        return values;
    }

    public void addValue(Object value) {
        values.add(value);
    }

    public void removeValue(Object value) {
        values.remove(value);
    }
}
