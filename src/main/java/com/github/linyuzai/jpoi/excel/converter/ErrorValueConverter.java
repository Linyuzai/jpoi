package com.github.linyuzai.jpoi.excel.converter;

public class ErrorValueConverter implements ValueConverter {

    private static ErrorValueConverter sInstance = new ErrorValueConverter();

    public static ErrorValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return false;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return null;
    }
}
