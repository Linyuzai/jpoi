package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.error.ByteError;
import com.github.linyuzai.jpoi.excel.value.error.IntError;

public class ErrorValueConverter implements ValueConverter {

    private static ErrorValueConverter sInstance = new ErrorValueConverter();

    public static ErrorValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value != null && (value.getClass() == byte.class || value.getClass() == int.class)
                || value instanceof Integer || value instanceof Byte;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if ((value != null && value.getClass() == byte.class) || value instanceof Byte) {
            return new ByteError((Byte) value);
        } else if ((value != null && value.getClass() == int.class) || value instanceof Integer) {
            return new IntError((Integer) value);
        }
        return null;
    }
}
