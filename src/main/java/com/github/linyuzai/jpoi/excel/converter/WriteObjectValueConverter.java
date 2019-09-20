package com.github.linyuzai.jpoi.excel.converter;

public class WriteObjectValueConverter implements ValueConverter {

    private static WriteObjectValueConverter sInstance = new WriteObjectValueConverter();

    public static WriteObjectValueConverter getInstance() {
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

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE;
    }
}
