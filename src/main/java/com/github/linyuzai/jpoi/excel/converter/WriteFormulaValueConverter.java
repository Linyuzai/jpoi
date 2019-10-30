package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.formula.FormulaValue;

public class WriteFormulaValueConverter implements ValueConverter {

    private static WriteFormulaValueConverter sInstance = new WriteFormulaValueConverter();

    public static WriteFormulaValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof String;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return new FormulaValue((String) value, null);
    }
}
