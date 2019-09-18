package com.github.linyuzai.jpoi.excel.write.converter;

public class ObjectValueConverter implements ValueConverter {

    private static ObjectValueConverter sInstance = new ObjectValueConverter();

    public static ObjectValueConverter getInstance() {
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
