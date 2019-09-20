package com.github.linyuzai.jpoi.excel.converter;

public class ReadObjectValueConverter implements ValueConverter {

    private static ReadObjectValueConverter sInstance = new ReadObjectValueConverter();

    public static ReadObjectValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return true;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return value;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE;
    }
}
