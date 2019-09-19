package com.github.linyuzai.jpoi.excel.read.adapter;

import java.util.HashMap;
import java.util.Map;

public class MapReadAdapter extends AnnotationReadAdapter<Map<String, Object>> {

    private Map<Integer, Map<Integer, ReadField>> readFieldMap = new HashMap<>();

    public void addField(int sheet, ReadField readField) {
        Map<Integer, ReadField> readFields = readFieldMap.computeIfAbsent(sheet, k -> new HashMap<>());
        readFields.put(readField.getIndex(), readField);
    }

    public void addField(String... fieldNames) {
        for (int i = 0; i < fieldNames.length; i++) {
            addField(i, fieldNames[i]);
        }
    }

    public void addField(int index, String fieldName) {
        ReadField readField = new ReadField();
        readField.setIndex(index);
        readField.setFieldName(fieldName);
        addField(0, readField);
    }


    @Override
    public Map<String, Object> createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        return null;
    }

    @Override
    public void adaptValue(Map<String, Object> cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }
}
