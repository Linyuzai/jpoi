package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.support.SupportValue;

public class WriteSupportValueConverter implements ValueConverter {

    private static WriteSupportValueConverter sInstance = new WriteSupportValueConverter();

    public static WriteSupportValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof SupportValue;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return value;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 1;
    }
}
