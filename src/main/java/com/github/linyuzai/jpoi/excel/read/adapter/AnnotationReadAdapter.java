package com.github.linyuzai.jpoi.excel.read.adapter;

import com.github.linyuzai.jpoi.excel.write.converter.ValueConverter;

public abstract class AnnotationReadAdapter extends ClassReadAdapter {

    private boolean annotationOnly;

    public boolean isAnnotationOnly() {
        return annotationOnly;
    }

    public void setAnnotationOnly(boolean annotationOnly) {
        this.annotationOnly = annotationOnly;
    }

    public static class AnnotationReadField extends ReadField {

        private boolean isMethod;
        private ValueConverter valueConverter;

        public boolean isMethod() {
            return isMethod;
        }

        public void setMethod(boolean method) {
            isMethod = method;
        }

        public ValueConverter getValueConverter() {
            return valueConverter;
        }

        public void setValueConverter(ValueConverter valueConverter) {
            this.valueConverter = valueConverter;
        }
    }
}
