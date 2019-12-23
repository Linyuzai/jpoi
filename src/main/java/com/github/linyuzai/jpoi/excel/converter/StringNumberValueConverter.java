package com.github.linyuzai.jpoi.excel.converter;

public class StringNumberValueConverter implements ValueConverter {

    private static final StringNumberValueConverter sInstance = new StringNumberValueConverter();

    public static StringNumberValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return true;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (value instanceof Double) {
            return String.valueOf(((Double) value).longValue());
        } else {
            return StringValueConverter.getInstance().convertValue(sheet, row, cell, value);
        }
    }
}
