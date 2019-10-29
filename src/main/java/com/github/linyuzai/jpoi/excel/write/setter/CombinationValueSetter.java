package com.github.linyuzai.jpoi.excel.write.setter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import org.apache.poi.ss.usermodel.*;

import java.util.Collection;

public class CombinationValueSetter extends SupportValueSetter {

    private static CombinationValueSetter sInstance = new CombinationValueSetter();

    public static CombinationValueSetter getInstance() {
        return sInstance;
    }

    @Override
    public void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper, Object value) {
        if (value instanceof CombinationValue) {
            setCombinationValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
            return;
        }
        super.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
    }

    public void setCombinationValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper, Object value) {
        if (value instanceof CombinationValue) {
            Object combinationValue = ((CombinationValue) value).getValue();
            if (combinationValue instanceof Collection) {
                for (Object o : (Collection) combinationValue) {
                    setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, o);
                }
            } else {
                setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, combinationValue);
            }
        }
    }
}
