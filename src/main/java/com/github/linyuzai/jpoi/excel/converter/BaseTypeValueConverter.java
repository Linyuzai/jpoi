package com.github.linyuzai.jpoi.excel.converter;

public class BaseTypeValueConverter implements ValueConverter {

    private static BaseTypeValueConverter sInstance = new BaseTypeValueConverter();

    public static BaseTypeValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value.getClass() == int.class || value instanceof Integer || value.getClass() == short.class || value instanceof Short
                || value.getClass() == byte.class || value instanceof Byte || value.getClass() == long.class || value instanceof Long
                || value.getClass() == float.class || value instanceof Float || value.getClass() == char.class || value instanceof Character;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (value.getClass() == int.class) {
            return Integer.valueOf((int) value).doubleValue();
        } else if (value instanceof Integer) {
            return ((Integer) value).doubleValue();
        } else if (value.getClass() == short.class) {
            return Short.valueOf((short) value).doubleValue();
        } else if (value instanceof Short) {
            return ((Short) value).doubleValue();
        } else if (value.getClass() == byte.class) {
            return Byte.valueOf((byte) value).doubleValue();
        } else if (value instanceof Byte) {
            return ((Byte) value).doubleValue();
        } else if (value.getClass() == long.class) {
            return Long.valueOf((long) value).doubleValue();
        } else if (value instanceof Long) {
            return ((Long) value).doubleValue();
        } else if (value.getClass() == float.class) {
            return Float.valueOf((float) value).doubleValue();
        } else if (value instanceof Float) {
            return ((Float) value).doubleValue();
        } else if (value.getClass() == char.class) {
            return String.valueOf(value);
        } else if (value instanceof Character) {
            return ((Character) value).toString();
        }
        return null;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 2;
    }
}
