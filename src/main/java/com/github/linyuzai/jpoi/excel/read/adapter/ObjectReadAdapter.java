package com.github.linyuzai.jpoi.excel.read.adapter;

import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.util.ClassUtils;

import java.lang.reflect.Field;
import java.util.*;

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
        combineFields(sheet);
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
            /*if (rf instanceof AnnotationReadField && (((AnnotationReadField) rf).isComment() || ((AnnotationReadField) rf).isPicture())) {
                continue;
            }*/
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

    public void combineFields(int sheet) {
        FieldData fieldData = getFieldDataMap().get(sheet);
        if (fieldData == null) {
            return;
        }
        Map<Integer, ReadField> readFieldMap = fieldData.getReadFieldMap();
        Set<Integer> keys = new HashSet<>(readFieldMap.keySet());
        for (Integer key : keys) {
            ReadField readField = readFieldMap.get(key);
            if (readField == null) {
                continue;
            }
            if (readField instanceof AnnotationReadField) {
                boolean isComment = ((AnnotationReadField) readField).isComment();
                boolean isPicture = ((AnnotationReadField) readField).isPicture();
                String commentField = ((AnnotationReadField) readField).getCommentOfField();
                int commentIndex = ((AnnotationReadField) readField).getCommentOfIndex();
                String pictureField = ((AnnotationReadField) readField).getPictureOfField();
                int pictureIndex = ((AnnotationReadField) readField).getPictureOfIndex();
                ValueConverter valueConverter = ((AnnotationReadField) readField).getValueConverter();
                //ValueConverter def = new DynamicValueConverter(ReadSupportValueConverter.getInstance(), ReadObjectValueConverter.getInstance());
                if (valueConverter == null) {
                    ((AnnotationReadField) readField).setValueConverter(DefaultReadValueConverter.getInstance());
                    if (isComment) {
                        ((AnnotationReadField) readField).setValueConverter(ReadCommentValueConverter.getInstance());
                    }
                    if (isPicture) {
                        ((AnnotationReadField) readField).setValueConverter(ReadPictureValueConverter.getInstance());
                    }
                }

                if (isComment && isPicture) {
                    throw new IllegalArgumentException("Can't set comments and pictures at the same time on the field '" + readField.getFieldName() + "'");
                } else if (isComment || isPicture) {
                    int index = commentIndex >= 0 ? commentIndex : pictureIndex;
                    if (index >= 0) {
                        if (readFieldMap.containsKey(index)) {
                            ReadField target = readFieldMap.get(index);
                            if (target instanceof AnnotationReadField) {
                                ((AnnotationReadField) target).getCombinationFields().add(readField);
                                if (((AnnotationReadField) target).getValueConverter() == null) {
                                    ((AnnotationReadField) target).setValueConverter(ReadDataValueConverter.getInstance());
                                }
                            }
                        }
                    } else {
                        String field = (commentField != null && !commentField.isEmpty()) ? commentField : pictureField;
                        for (ReadField rf : readFieldMap.values()) {
                            if (rf instanceof AnnotationReadField && field.equals(rf.getFieldName())) {
                                ((AnnotationReadField) rf).getCombinationFields().add(readField);
                                if (((AnnotationReadField) rf).getValueConverter() == null) {
                                    ((AnnotationReadField) rf).setValueConverter(ReadDataValueConverter.getInstance());
                                }
                                break;
                            }
                        }
                    }
                    readFieldMap.remove(key);
                } else {
                    //((AnnotationReadField) readField).setValueConverter(DefaultReadValueConverter.getInstance());
                }
            }
        }
        fieldData.setReadFields(new ArrayList<>(readFieldMap.values()));
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
                //List<ReadField> rfs = new ArrayList<>(fieldData.getReadFieldMap().values());
                //ReadField readField = rfs.get(c);
                ReadField readField = fieldData.getReadFields().get(c);
                if (readField != null) {
                    setFieldValue(cellContainer, value, s, r, c, readField);
                    if (value instanceof CombinationValue ||
                            readField instanceof AnnotationReadField && ((AnnotationReadField) readField).getCombinationFields().size() > 0) {
                        for (ReadField combinationField : ((AnnotationReadField) readField).getCombinationFields()) {
                            setFieldValue(cellContainer, value, s, r, c, combinationField);
                        }
                    }
                    /*if (value instanceof CombinationValue) {
                        setFieldValue(cellContainer, value, s, r, c, readField);
                        if (readField instanceof AnnotationReadField && ((AnnotationReadField) readField).getCombinationFields().size() > 0) {
                            for (ReadField combinationField : ((AnnotationReadField) readField).getCombinationFields()) {
                                setFieldValue(cellContainer, value, s, r, c, combinationField);
                            }
                        }
                    } else {
                        setFieldValue(cellContainer, value, s, r, c, readField);
                    }*/
                }
            }
        }
    }

    public void setFieldValue(Object cellContainer, Object value, int s, int r, int c, ReadField readField) {
        ValueConverter valueConverter;
        Object val = value;
        String fieldName = readField.getFieldName();
        if (fieldName != null) {
            try {
                if (readField instanceof AnnotationReadField &&
                        (valueConverter = ((AnnotationReadField) readField).getValueConverter()) != null &&
                        valueConverter.supportValue(s, r, c, value)) {
                    val = valueConverter.convertValue(s, r, c, value);
                }
                Field field = classes[s].getDeclaredField(fieldName);
                field.setAccessible(true);
                field.set(cellContainer, baseConvert(field.getType(), val));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private Object baseConvert(Class<?> cls, Object value) {
        if (value == null || value.getClass() == cls) {
            return value;
        }
        if (cls.isPrimitive() || ClassUtils.isWrapClass(cls)) {
            if (cls == int.class || cls == Integer.class) {
                return Double.valueOf(value.toString()).intValue();
            } else if (cls == short.class || cls == Short.class) {
                return Double.valueOf(value.toString()).shortValue();
            } else if (cls == byte.class || cls == Byte.class) {
                return Double.valueOf(value.toString()).byteValue();
            } else if (cls == long.class || cls == Long.class) {
                return Double.valueOf(value.toString()).longValue();
            } else if (cls == float.class || cls == Float.class) {
                return Double.valueOf(value.toString()).floatValue();
            } else if (cls == char.class || cls == Character.class) {
                return value.toString().charAt(0);
            }
        }
        return value;
    }
}
