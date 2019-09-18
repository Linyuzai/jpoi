package com.github.linyuzai.jpoi.excel.read.adapter;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class ObjectReadAdapter extends MapReadAdapter {

    private boolean toMap;

    public void setTargetClass(Class<?> cls) {
        ReadField readField = getReadFieldIncludeAnnotation(cls);
        if (readField instanceof AnnotationReadField){

        }
    }

    @Override
    public Object readDataRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook) {
        if (toMap) {
            return super.readDataRow(cellValues, s, r, row, sheet, workbook);
        } else {

            return null;
        }
    }

    public boolean isToMap() {
        return toMap;
    }

    public void setToMap(boolean toMap) {
        this.toMap = toMap;
    }
}
