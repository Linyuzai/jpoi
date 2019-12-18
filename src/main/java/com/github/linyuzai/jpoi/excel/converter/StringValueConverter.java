package com.github.linyuzai.jpoi.excel.converter;

public class StringValueConverter implements ValueConverter {

    private static final StringValueConverter sInstance = new StringValueConverter();

    public static StringValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return true;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return String.valueOf(value);
    }
}
