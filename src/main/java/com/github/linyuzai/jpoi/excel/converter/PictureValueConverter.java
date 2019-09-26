package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.picture.*;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;

import java.awt.image.BufferedImage;
import java.io.File;

public class PictureValueConverter extends ClientAnchorValueConverter {

    private static PictureValueConverter sInstance = new PictureValueConverter();

    public static PictureValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public boolean supportValue(int sheet, int row, int cell, Object value) {
        return value instanceof File || value instanceof BufferedImage || value instanceof byte[] || value instanceof String;
    }

    @Override
    public Object convertValue(int sheet, int row, int cell, Object value) {
        PoiPicture picture = null;
        if (value instanceof byte[]) {
            picture = new ByteArrayPicture((byte[]) value);
        } else if (value instanceof BufferedImage) {
            picture = new BufferedImagePicture((BufferedImage) value);
        } else if (value instanceof File) {
            picture = new FilePicture((File) value);
        } else if (value instanceof String) {
            picture = new Base64Picture((String) value);
        }
        if (picture != null) {
            configClientAnchorValue(sheet, row, cell, value, picture);
            picture.setType(getType(sheet, row, cell, value));
            return picture;
        }
        return null;
    }

    public int getType(int sheet, int row, int cell, Object value) {
        return Workbook.PICTURE_TYPE_JPEG;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 2;
    }
}
