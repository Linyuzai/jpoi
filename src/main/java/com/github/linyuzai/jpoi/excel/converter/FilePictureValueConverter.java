package com.github.linyuzai.jpoi.excel.converter;

import java.io.File;
import java.net.URI;

public class FilePictureValueConverter extends WritePictureValueConverter {

    private static FilePictureValueConverter sInstance = new FilePictureValueConverter();

    public static FilePictureValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof File || value instanceof String || value instanceof URI;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        if (value instanceof String) {
            return super.convertValue(sheet, row, cell, new File((String) value));
        } else if (value instanceof URI) {
            return super.convertValue(sheet, row, cell, new File((URI) value));
        }
        return super.convertValue(sheet, row, cell, value);
    }
}
