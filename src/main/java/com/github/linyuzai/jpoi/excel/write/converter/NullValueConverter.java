package com.github.linyuzai.jpoi.excel.write.converter;

public class NullValueConverter implements ValueConverter {

    private static NullValueConverter sInstance = new NullValueConverter();

    public static NullValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value == null;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return null;
    }

    @Override
    public int getOrder() {
        return Integer.MIN_VALUE;
    }
}
