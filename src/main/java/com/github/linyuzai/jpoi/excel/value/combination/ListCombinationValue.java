package com.github.linyuzai.jpoi.excel.value.combination;

import java.util.ArrayList;
import java.util.List;

public class ListCombinationValue implements CombinationValue {

    private List<Object> values = new ArrayList<>();

    @Override
    public Object getValue() {
        return values;
    }

    @Override
    public void setValue(Object value) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void addValue(Object value) {
        values.add(value);
    }

    @Override
    public void removeValue(Object value) {
        values.remove(value);
    }
}
