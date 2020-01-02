package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.post.PostValue;

public class PostValueConverter implements ValueConverter {

    private static PostValueConverter sInstance = new PostValueConverter();

    public static PostValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof PostValue;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        return value;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 2;
    }
}
