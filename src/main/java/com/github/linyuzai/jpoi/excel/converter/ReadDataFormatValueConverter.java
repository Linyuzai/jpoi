package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.comment.SupportComment;
import com.github.linyuzai.jpoi.excel.value.error.SupportErrorValue;
import com.github.linyuzai.jpoi.excel.value.format.SupportDataFormat;
import com.github.linyuzai.jpoi.excel.value.formula.SupportFormula;
import com.github.linyuzai.jpoi.excel.value.picture.*;
import com.github.linyuzai.jpoi.support.SupportValue;

public class ReadDataFormatValueConverter implements ValueConverter {

    private static ReadDataFormatValueConverter sInstance = new ReadDataFormatValueConverter();

    public static ReadDataFormatValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof SupportDataFormat;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (value instanceof SupportDataFormat) {
            return ((SupportDataFormat) value).getFormatValue();
        }
        return null;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 4;
    }
}
