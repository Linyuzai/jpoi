package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.SupportValue;

public class SupportValueConverter implements ValueConverter {

    private static SupportValueConverter sInstance = new SupportValueConverter();

    public static SupportValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof SupportValue;
    }

    @Override
    public Object adaptValue(int sheet, int row, int cell, Object value) {
        return value;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 1;
    }
}
