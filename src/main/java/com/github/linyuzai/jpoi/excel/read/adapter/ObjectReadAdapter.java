package com.github.linyuzai.jpoi.excel.read.adapter;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

public class ObjectReadAdapter extends MapReadAdapter {

    private Class<?>[] classes;

    public ObjectReadAdapter(Class<?>... classes) {
        setTargetClass(classes);
    }

    public void setTargetClass(Class<?>... classes) {
        this.classes = classes;
        for (int i = 0; i < classes.length; i++) {
            setTargetClass(i, classes[i]);
        }
    }

    private void setTargetClass(int sheet, Class<?> cls) {
        if (Map.class.isAssignableFrom(cls)) {
            return;
        }
        fieldsFromClass(sheet, cls);
    }

    public void fieldsFromClass(int sheet, Class<?> cls) {
        ReadField crf = getReadFieldIncludeAnnotation(cls);
        boolean toMap = false;
        boolean annotationOnly = false;
        String sheetName = null;
        if (crf instanceof AnnotationReadField) {
            toMap = ((AnnotationReadField) crf).isToMap();
            annotationOnly = ((AnnotationReadField) crf).isAnnotationOnly();
            sheetName = crf.getFieldDescription();
        }
        List<ReadField> rfs = new ArrayList<>();
        List<Integer> indexes = new ArrayList<>();
        Field[] fields = cls.getDeclaredFields();
        for (int f = 0; f < fields.length; f++) {
            ReadField frf = getReadFieldIncludeAnnotation(fields[f]);
            if (annotationOnly) {
                if (frf instanceof AnnotationReadField) {
                    ((AnnotationReadField) frf).setToMap(toMap);
                    ((AnnotationReadField) frf).setAnnotationOnly(true);
                    rfs.add(frf);
                }
            } else {
                rfs.add(frf);
            }
            indexes.add(f);
        }
        for (ReadField rf : rfs) {
            if (rf.getIndex() > -1) {
                indexes.remove(rf.getIndex());
            }
        }
        int i = 0;
        for (ReadField rf : rfs) {
            if (rf.getIndex() < 0) {
                int last = indexes.get(indexes.size() - 1);
                rf.setIndex(i < indexes.size() ? indexes.get(i) : last + 1);
            }
            i++;
        }
        rfs.sort(Comparator.comparingInt(ReadField::getIndex));
        for (ReadField rf : rfs) {
            addFieldData(sheet, rf.getIndex(), rf);
        }
    }

    @Override
    public void readRowHeaderCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        if (Map.class.isAssignableFrom(classes[s])) {
            super.readRowHeaderCell(value, s, r, c, sCount, rCount, cCount);
        }
    }

    @Override
    public Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        FieldData fieldData = getFieldDataMap().get(s);
        boolean toMap = fieldData != null && fieldData.isToMap();
        if (toMap) {
            return super.createContainer(value, s, r, c, sCount, rCount, cCount);
        }
        try {
            return classes[s].newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        FieldData fieldData = getFieldDataMap().get(s);
        boolean toMap = fieldData != null && fieldData.isToMap();
        if (toMap) {
            super.adaptValue(cellContainer, value, s, r, c, sCount, rCount, cCount);
        } else {
            if (fieldData != null) {
                ReadField readField = fieldData.getReadFieldMap().get(c);
                if (readField != null) {
                    String fieldName = readField.getFieldName();
                    if (fieldName != null) {
                        try {
                            Field field = classes[s].getDeclaredField(fieldName);
                            field.setAccessible(true);
                            field.set(cellContainer, value);
                        } catch (NoSuchFieldException | IllegalAccessException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        }
    }
}
