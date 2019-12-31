package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

public class WriteUrlPictureValueConverter implements ValueConverter {

    private static WriteUrlPictureValueConverter sInstance = new WriteUrlPictureValueConverter();

    public static WriteUrlPictureValueConverter getInstance() {
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
        holder.setValueConverter(WritePictureValueConverter.getInstance());
        return holder;
    }

}
