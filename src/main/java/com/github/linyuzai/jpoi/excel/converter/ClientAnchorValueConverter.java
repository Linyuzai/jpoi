package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.ClientAnchorValue;
import com.github.linyuzai.jpoi.excel.value.picture.PoiPicture;
import org.apache.poi.ss.usermodel.ClientAnchor;

public abstract class ClientAnchorValueConverter implements ValueConverter {

    public void configClientAnchorValue(int sheet, int row, int cell,Object value, ClientAnchorValue clientAnchorValue) {
        if (clientAnchorValue != null) {
            clientAnchorValue.setLocation(getLocation(sheet, row, cell, value));
            clientAnchorValue.setPadding(getPadding(sheet, row, cell, value));
            clientAnchorValue.setAnchorType(getAnchorType(sheet, row, cell, value));
        }
    }

    public PoiPicture.Location getLocation(int sheet, int row, int cell, Object value) {
        return new PoiPicture.Location(row, cell, row + 1, cell + 1);
    }

    public PoiPicture.Padding getPadding(int sheet, int row, int cell, Object value) {
        return PoiPicture.Padding.getDefault();
    }

    public ClientAnchor.AnchorType getAnchorType(int sheet, int row, int cell, Object value) {
        return ClientAnchor.AnchorType.MOVE_AND_RESIZE;
    }
}
