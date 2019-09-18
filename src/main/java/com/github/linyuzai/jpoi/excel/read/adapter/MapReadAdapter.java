package com.github.linyuzai.jpoi.excel.read.adapter;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MapReadAdapter extends AnnotationReadAdapter {

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
    public Object readDataRow(List<?> cellValues, int s, int r, Row row, Sheet sheet, Workbook workbook) {
        Map<String, Object> valueMap = new HashMap<>();
        for (int i = 0; i < cellValues.size(); i++) {
            if (readFieldMap.containsKey(i)) {
                valueMap.put(readFieldMap.get(s).get(i).getFieldName(), cellValues.get(i));
            }
        }
        return valueMap;
    }
}
