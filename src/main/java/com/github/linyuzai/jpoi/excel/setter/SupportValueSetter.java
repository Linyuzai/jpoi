package com.github.linyuzai.jpoi.excel.setter;

import com.github.linyuzai.jpoi.excel.value.ByteArrayPicture;
import com.github.linyuzai.jpoi.excel.value.SupportPicture;
import com.github.linyuzai.jpoi.excel.value.SupportValue;
import org.apache.poi.ss.usermodel.*;

public class SupportValueSetter extends PoiValueSetter {

    @Override
    public void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, Object value) {
        if (value instanceof SupportValue) {
            setSupportValue(s, r, c, cell, row, sheet, drawing, workbook, value);
            return;
        }
        super.setValue(s, r, c, cell, row, sheet, drawing, workbook, value);
    }

    public void setSupportValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, Object value) {
        if (value instanceof SupportPicture) {
            SupportPicture.Location location = ((SupportPicture) value).getLocation();
            SupportPicture.Padding padding = ((SupportPicture) value).getPadding();
            int type = ((SupportPicture) value).getType();
            ClientAnchor anchor = drawing.createAnchor(
                    padding.getLeft(), padding.getTop(), padding.getRight(), padding.getBottom(),
                    location.getStartCell(), location.getStartRow(), location.getEndCell(), location.getEndRow());
            if (value instanceof ByteArrayPicture) {
                drawing.createPicture(anchor, workbook.addPicture(((ByteArrayPicture) value).getBytes(), type));
            }
        }
    }
}
