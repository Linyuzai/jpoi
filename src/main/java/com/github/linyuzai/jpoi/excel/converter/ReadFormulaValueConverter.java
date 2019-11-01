package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.excel.value.formula.SupportFormula;

import java.util.Collection;

public class ReadFormulaValueConverter implements ValueConverter {

    private static ReadFormulaValueConverter sInstance = new ReadFormulaValueConverter();

    public static ReadFormulaValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof SupportFormula || value instanceof CombinationValue;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (((CombinationValue) value).getValue(null) instanceof Collection) {
            for (Object o : ((Collection) ((CombinationValue) value).getValue(null))) {
                if (o instanceof SupportFormula) {
                    return getFormula(sheet, row, cell, (SupportFormula) o);
                }
            }
        } else {
            if (value instanceof SupportFormula) {
                return getFormula(sheet, row, cell, (SupportFormula) value);
            } else if (((CombinationValue) value).getValue(null) instanceof SupportFormula) {
                return getFormula(sheet, row, cell, (SupportFormula) ((CombinationValue) value).getValue(null));
            }
        }
        return null;
    }

    public Object getFormula(int sheet, int row, int cell, SupportFormula formula) {
        Object value = ReadSupportValueConverter.getInstance().convertValue(sheet, row, cell, formula);
        if (value instanceof SupportFormula) {
            return ((SupportFormula) value).getValue();
        } else {
            return value;
        }
    }
}
