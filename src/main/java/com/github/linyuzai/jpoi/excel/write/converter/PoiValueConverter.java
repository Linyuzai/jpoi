package com.github.linyuzai.jpoi.excel.write.converter;

import org.apache.poi.ss.usermodel.RichTextString;

import java.util.Calendar;
import java.util.Date;

public class PoiValueConverter implements ValueConverter {

    private static PoiValueConverter sInstance = new PoiValueConverter();

    public static PoiValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value == null || value.getClass() == boolean.class || value.getClass() == double.class || value instanceof Boolean
                || value instanceof Double || value instanceof String || value instanceof Date || value instanceof Calendar
                || value instanceof RichTextString;
    }

    @Override
    public Object adaptValue(int sheet, int row, int cell, Object value) {
        return value;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 2;
    }
}
