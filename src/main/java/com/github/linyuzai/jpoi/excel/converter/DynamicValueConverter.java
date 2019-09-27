package com.github.linyuzai.jpoi.excel.converter;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class DynamicValueConverter implements ValueConverter {

    private List<ValueConverter> valueConverters = new ArrayList<>();

    public DynamicValueConverter() {
    }

    public DynamicValueConverter(ValueConverter... valueConverters) {
        this.valueConverters.addAll(Arrays.asList(valueConverters));
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public void setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        for (ValueConverter valueConverter : valueConverters) {
            if (valueConverter.supportValue(sheet, row, cell, value)) {
                return true;
            }
        }
        return false;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        for (ValueConverter valueConverter : valueConverters) {
            if (valueConverter.supportValue(sheet, row, cell, value)) {
                return valueConverter.convertValue(sheet, row, cell, value);
            }
        }
        return null;
    }
}
