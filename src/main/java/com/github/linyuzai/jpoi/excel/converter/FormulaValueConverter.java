package com.github.linyuzai.jpoi.excel.converter;

public class FormulaValueConverter implements ValueConverter {

    private static FormulaValueConverter sInstance = new FormulaValueConverter();

    public static FormulaValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof String;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return null;
    }
}
