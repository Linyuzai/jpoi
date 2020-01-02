package com.github.linyuzai.jpoi.excel.read.adapter;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.util.HashMap;
import java.util.List;
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
    public Object readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        String str = String.valueOf(value);
        addFieldData(s, c, str, str);
        return str;
    }

    @Override
    public Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return new HashMap<String, Object>();
    }

    @SuppressWarnings("unchecked")
    @Override
    public Object adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable {
        FieldData fieldData = fieldDataMap.get(s);
        if (fieldData != null) {
            ReadField readField = fieldData.getReadFieldMap().get(c);
            if (readField != null && readField.getFieldName() != null) {
                ValueConverter valueConverter;
                Object val = value;
                if (readField instanceof AnnotationReadField &&
                        (valueConverter = ((AnnotationReadField) readField).getValueConverter()) != null) {
                    val = valueConverter.convertValue(s, r, c, value);
                }
                ((Map<String, Object>) cellContainer).put(readField.getFieldName(), val);
                return val;
            }
        }
        return null;
    }

    @Override
    public int getHeaderRowCount(int sheet) {
        return 1;
    }

    public static class FieldData {
        private String sheetName;
        private boolean toMap;
        private Map<Integer, ReadField> readFieldMap;
        private List<ReadField> readFields;

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

        public List<ReadField> getReadFields() {
            return readFields;
        }

        public void setReadFields(List<ReadField> readFields) {
            this.readFields = readFields;
        }
    }
}
