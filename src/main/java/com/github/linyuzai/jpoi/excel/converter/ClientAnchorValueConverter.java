package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.excel.value.ClientAnchorValue;
import com.github.linyuzai.jpoi.excel.value.SupportClientAnchorValue;
import org.apache.poi.ss.usermodel.ClientAnchor;

public abstract class ClientAnchorValueConverter implements ValueConverter {

    public void configClientAnchorValue(int sheet, int row, int cell,Object value, ClientAnchorValue clientAnchorValue) {
        if (clientAnchorValue != null) {
            clientAnchorValue.setLocation(getLocation(sheet, row, cell, value));
            clientAnchorValue.setPadding(getPadding(sheet, row, cell, value));
            clientAnchorValue.setAnchorType(getAnchorType(sheet, row, cell, value));
        }
    }

    public SupportClientAnchorValue.Location getLocation(int sheet, int row, int cell, Object value) {
        return new SupportClientAnchorValue.Location(row, cell, row + 1, cell + 1);
    }

    public SupportClientAnchorValue.Padding getPadding(int sheet, int row, int cell, Object value) {
        return SupportClientAnchorValue.Padding.getDefault();
    }

    public ClientAnchor.AnchorType getAnchorType(int sheet, int row, int cell, Object value) {
        return ClientAnchor.AnchorType.MOVE_AND_RESIZE;
    }
}
