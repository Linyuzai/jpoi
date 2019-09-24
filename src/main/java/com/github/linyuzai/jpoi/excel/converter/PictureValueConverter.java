package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.picture.*;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;

import java.awt.image.BufferedImage;
import java.io.File;

public class PictureValueConverter implements ValueConverter {

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
            picture.setLocation(getLocation(sheet, row, cell, value));
            picture.setPadding(getPadding(sheet, row, cell, value));
            picture.setAnchorType(getAnchorType(sheet, row, cell, value));
            picture.setType(getType(sheet, row, cell, value));
            return picture;
        }
        return null;
    }

    public PoiPicture.Location getLocation(int sheet, int row, int cell, Object value) {
        return new PoiPicture.Location(row, cell, row + 1, cell + 1);
    }

    public PoiPicture.Padding getPadding(int sheet, int row, int cell, Object value) {
        return PoiPicture.Padding.getDefault();
    }

    public int getType(int sheet, int row, int cell, Object value) {
        return Workbook.PICTURE_TYPE_JPEG;
    }

    public ClientAnchor.AnchorType getAnchorType(int sheet, int row, int cell, Object value) {
        return ClientAnchor.AnchorType.MOVE_AND_RESIZE;
    }

    @Override
    public int getOrder() {
        return Integer.MAX_VALUE - 2;
    }
}
