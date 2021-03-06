package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.excel.value.picture.SupportPicture;

import java.util.Collection;

public class ReadPictureValueConverter implements ValueConverter {

    private static ReadPictureValueConverter sInstance = new ReadPictureValueConverter();

    public static ReadPictureValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof CombinationValue || value instanceof SupportPicture;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (((CombinationValue) value).getValue(null) instanceof Collection) {
            for (Object o : ((Collection) ((CombinationValue) value).getValue(null))) {
                if (o instanceof SupportPicture) {
                    return getPicture(sheet, row, cell, (SupportPicture) o);
                }
            }
        } else {
            if (value instanceof SupportPicture) {
                return getPicture(sheet, row, cell, (SupportPicture) value);
            } else if (((CombinationValue) value).getValue(null) instanceof SupportPicture) {
                return getPicture(sheet, row, cell, (SupportPicture) ((CombinationValue) value).getValue(null));
            }
        }
        return null;
    }

    public Object getPicture(int sheet, int row, int cell, SupportPicture picture) {
        return ReadSupportValueConverter.getInstance().convertValue(sheet, row, cell, picture);
    }
}
