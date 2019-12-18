package com.github.linyuzai.jpoi.excel.write.setter;

import com.github.linyuzai.jpoi.excel.converter.BaseTypeValueConverter;
import org.apache.poi.ss.usermodel.*;

public class BaseTypeValueSetter extends PoiValueSetter {

    @Override
    public void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper, Object value) throws Throwable {
        if (BaseTypeValueConverter.getInstance().supportValue(s, r, c, value)) {
            super.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, BaseTypeValueConverter.getInstance().convertValue(s, r, c, value));
        } else {
            super.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
        }
    }
}
