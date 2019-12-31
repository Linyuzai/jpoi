package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

public class ReadUrlPictureValueConverter implements ValueConverter {

    private static ReadUrlPictureValueConverter sInstance = new ReadUrlPictureValueConverter();

    public static ReadUrlPictureValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof String;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        UrlPostValueHolder holder = new UrlPostValueHolder();
        holder.setSource((String) value);
        return holder;
    }
}
