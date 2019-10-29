package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;

public class WriteCombinationValueConverter implements ValueConverter {

    private static WriteCombinationValueConverter sInstance = new WriteCombinationValueConverter();

    public static WriteCombinationValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof CombinationValue;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return value;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 2;
    }
}
