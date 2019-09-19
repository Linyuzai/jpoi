package com.github.linyuzai.jpoi.excel.read.adapter;

import com.github.linyuzai.jpoi.excel.read.annotation.JExcelReadSheet;
import com.github.linyuzai.jpoi.excel.write.converter.ValueConverter;

public abstract class AnnotationReadAdapter extends ClassReadAdapter {

    private boolean annotationOnly;

    public ReadField getReadFieldIncludeAnnotation(Class<?> cls) {
        if (cls == null) {
            return null;
        } else {
            JExcelReadSheet ca = cls.getAnnotation(JExcelReadSheet.class);
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

    public boolean isAnnotationOnly() {
        return annotationOnly;
    }

    public void setAnnotationOnly(boolean annotationOnly) {
        this.annotationOnly = annotationOnly;
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
