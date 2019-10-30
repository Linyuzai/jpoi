package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;

public class DefaultReadValueConverter extends ReadDataValueConverter {

    private static DefaultReadValueConverter sInstance = new DefaultReadValueConverter();

    public static DefaultReadValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return true;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (value instanceof CombinationValue) {
            return super.convertValue(sheet, row, cell, value);
        } else {
            return value;
        }
    }
}
