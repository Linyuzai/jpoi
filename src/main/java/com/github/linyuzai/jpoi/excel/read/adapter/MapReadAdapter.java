package com.github.linyuzai.jpoi.excel.read.adapter;

import java.util.HashMap;
import java.util.Map;

public class MapReadAdapter extends AnnotationReadAdapter {

    private Map<Integer, FieldData> fieldDataMap = new HashMap<>();

    public Map<Integer, FieldData> getFieldDataMap() {
        return fieldDataMap;
    }

    public void addFieldData(int sheet, String... fieldDescriptions) {
        for (int i = 0; i < fieldDescriptions.length; i++) {
            addFieldData(sheet, i, null, fieldDescriptions[i]);
        }
    }

    public void addFieldData(int sheet, int index, String fieldName, String fieldDescription) {
        ReadField readField = new ReadField();
        readField.setIndex(index);
        readField.setFieldName(fieldName);
        readField.setFieldDescription(fieldDescription);
        addFieldData(sheet, index, readField);
    }

    public void addFieldData(int sheet, int index, ReadField readField) {
        FieldData fieldData = fieldDataMap.get(sheet);
        if (fieldData == null) {
            fieldData = new FieldData();
            fieldData.setReadFieldMap(new HashMap<>());
            addFieldData(sheet, fieldData);
        }
        fieldData.getReadFieldMap().put(index, readField);
    }

    public void addFieldData(int sheet, FieldData fieldData) {
        fieldDataMap.put(sheet, fieldData);
    }

    @Override
    public void readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        String str = String.valueOf(value);
        addFieldData(s, c, str, str);
    }

    @Override
    public Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return new HashMap<String, Object>();
    }

    @SuppressWarnings("unchecked")
    @Override
    public void adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        FieldData fieldData = fieldDataMap.get(s);
        if (fieldData != null) {
            ReadField readField = fieldData.getReadFieldMap().get(c);
            if (readField != null && readField.getFieldName() != null) {
                ((Map<String, Object>) cellContainer).put(readField.getFieldName(), value);
            }
        }
    }

    @Override
    public int getHeaderRowCount(int sheet) {
        return 1;
    }

    public static class FieldData {
        private String sheetName;
        private boolean toMap;
        private Map<Integer, ReadField> readFieldMap;

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public boolean isToMap() {
            return toMap;
        }

        public void setToMap(boolean toMap) {
            this.toMap = toMap;
        }

        public Map<Integer, ReadField> getReadFieldMap() {
            return readFieldMap;
        }

        public void setReadFieldMap(Map<Integer, ReadField> readFieldMap) {
            this.readFieldMap = readFieldMap;
        }
    }
}