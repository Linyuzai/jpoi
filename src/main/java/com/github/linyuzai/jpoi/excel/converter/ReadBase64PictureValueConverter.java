package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.picture.SupportPicture;

import java.util.Base64;

public class ReadBase64PictureValueConverter extends ReadPictureValueConverter {

    private static ReadBase64PictureValueConverter sInstance = new ReadBase64PictureValueConverter();

    public static ReadBase64PictureValueConverter getInstance() {
        return sInstance;
    }

    @Override
    public Object getPicture(int sheet, int row, int cell, SupportPicture picture) {
        return Base64.getEncoder().encodeToString((byte[]) super.getPicture(sheet, row, cell, picture));
    }
}
