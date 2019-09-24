package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.error.SupportErrorValue;
import com.github.linyuzai.jpoi.excel.value.formula.SupportFormula;
import com.github.linyuzai.jpoi.excel.value.picture.*;
import com.github.linyuzai.jpoi.support.SupportValue;

public class ReadSupportValueConverter implements ValueConverter {

    private static ReadSupportValueConverter sInstance = new ReadSupportValueConverter();

    public static ReadSupportValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof SupportValue;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (value instanceof SupportPicture) {
            if (value instanceof ByteArrayPicture) {
                return ((ByteArrayPicture) value).getBytes();
            } else if (value instanceof BufferedImagePicture) {
                return ((BufferedImagePicture) value).getBufferedImage();
            } else if (value instanceof FilePicture) {
                return ((FilePicture) value).getFile();
            } else if (value instanceof Base64Picture) {
                return ((Base64Picture) value).getBase64();
            }
        } else if (value instanceof SupportErrorValue) {
            return ((SupportErrorValue) value).getErrorValue();
        } else if (value instanceof SupportFormula) {
            return ((SupportFormula) value).getFormula();
        }
        return null;
    }
}
