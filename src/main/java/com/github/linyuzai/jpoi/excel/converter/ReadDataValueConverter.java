package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.excel.value.comment.SupportComment;
import com.github.linyuzai.jpoi.excel.value.picture.SupportPicture;
import com.github.linyuzai.jpoi.support.SupportValue;

import java.util.Collection;

public class ReadDataValueConverter implements ValueConverter {

    private static ReadDataValueConverter sInstance = new ReadDataValueConverter();

    public static ReadDataValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof CombinationValue;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        Object cValue = ((CombinationValue) value).getValue(null);
        if (cValue instanceof Collection) {
            for (Object o : ((Collection) cValue)) {
                if (!(o instanceof SupportPicture) && !(o instanceof SupportComment)) {
                    return getData(sheet, row, cell, o);
                }
            }
        } else {
            return getData(sheet, row, cell, ((CombinationValue) value).getValue(null));
        }
        return null;
    }

    public Object getData(int sheet, int row, int cell, Object value) {
        if (value instanceof SupportValue) {
            return ReadSupportValueConverter.getInstance().convertValue(sheet, row, cell, value);
        }
        return ReadObjectValueConverter.getInstance().convertValue(sheet, row, cell, value);
    }
}
