package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.comment.StringComment;

public class CommentValueConverter extends ClientAnchorValueConverter {

    private static CommentValueConverter sInstance = new CommentValueConverter();

    public static CommentValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof String;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        StringComment comment = new StringComment((String) value);
        configClientAnchorValue(sheet, row, cell, value, comment);
        return comment;
    }
}
