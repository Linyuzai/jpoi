package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.excel.value.comment.SupportComment;

import java.util.Collection;

public class ReadCommentValueConverter implements ValueConverter {

    private static ReadCommentValueConverter sInstance = new ReadCommentValueConverter();

    public static ReadCommentValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof CombinationValue || value instanceof SupportComment;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (((CombinationValue) value).getValue(null) instanceof Collection) {
            for (Object o : ((Collection) ((CombinationValue) value).getValue(null))) {
                if (o instanceof SupportComment) {
                    return getComment(sheet, row, cell, (SupportComment) o);
                }
            }
        } else {
            if (value instanceof SupportComment) {
                return getComment(sheet, row, cell, (SupportComment) value);
            } else if (((CombinationValue) value).getValue(null) instanceof SupportComment) {
                return getComment(sheet, row, cell, (SupportComment) ((CombinationValue) value).getValue(null));
            }
        }
        return null;
    }

    public Object getComment(int sheet, int row, int cell, SupportComment picture) {
        return ReadSupportValueConverter.getInstance().convertValue(sheet, row, cell, picture);
    }
}
