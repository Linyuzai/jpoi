package com.github.linyuzai.jpoi.excel.read.adapter;

import com.github.linyuzai.jpoi.excel.read.annotation.JExcelCellReader;
import com.github.linyuzai.jpoi.excel.read.annotation.JExcelSheetReader;
import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.lang.reflect.Field;

public abstract class AnnotationReadAdapter extends ClassReadAdapter {

    public ReadField getReadFieldIncludeAnnotation(Class<?> cls) {
        if (cls == null) {
            return null;
        } else {
            JExcelSheetReader ca = cls.getAnnotation(JExcelSheetReader.class);
            if (ca == null) {
                return null;
            }
            AnnotationReadField readField = new AnnotationReadField();
            readField.setToMap(ca.toMap());
            readField.setFieldName(cls.getName());
            readField.setAnnotationOnly(ca.annotationOnly());
            String title;
            if ((title = ca.name().trim()).isEmpty()) {
                readField.setFieldDescription(null);
            } else {
                readField.setFieldDescription(title);
            }
            return readField;
        }
    }

    public ReadField getReadFieldIncludeAnnotation(Field field) {
        JExcelCellReader fa = field.getAnnotation(JExcelCellReader.class);
        if (fa == null) {
            ReadField readField = new ReadField();
            readField.setFieldName(field.getName());
            readField.setFieldDescription(field.getName());
            readField.setIndex(-1);
            //readField.setOrder(Integer.MAX_VALUE);
            return readField;
        } else {
            String title = fa.title().trim();
            AnnotationReadField readField = new AnnotationReadField();
            readField.setFieldName(field.getName());
            readField.setFieldDescription(title.isEmpty() ? field.getName() : title);
            readField.setIndex(fa.index());
            //readField.setOrder(fa.order());
            //reuseValueConverter(writeField, fa.valueConverter());
            return readField;
        }
    }

    public static class AnnotationReadField extends ReadField {

        private boolean toMap;
        private boolean annotationOnly;
        private ValueConverter valueConverter;

        public boolean isToMap() {
            return toMap;
        }

        public void setToMap(boolean toMap) {
            this.toMap = toMap;
        }

        public boolean isAnnotationOnly() {
            return annotationOnly;
        }

        public void setAnnotationOnly(boolean annotationOnly) {
            this.annotationOnly = annotationOnly;
        }

        public ValueConverter getValueConverter() {
            return valueConverter;
        }

        public void setValueConverter(ValueConverter valueConverter) {
            this.valueConverter = valueConverter;
        }
    }
}
