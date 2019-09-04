package com.github.linyuzai.jpoi.excel.adapter;

import com.github.linyuzai.jpoi.excel.annotation.JExcelCell;
import com.github.linyuzai.jpoi.excel.annotation.JExcelSheet;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public abstract class AnnotationWriteAdapter extends ClassWriteAdapter {

    private boolean annotationOnly;

    public FieldData getFieldDataIncludeAnnotation(Class<?> cls) {
        if (cls == null) {
            return null;
        } else {
            JExcelSheet ca = cls.getAnnotation(JExcelSheet.class);
            String title;
            if (ca == null || (title = ca.name().trim()).isEmpty()) {
                return null;
            } else {
                AnnotationFieldData fieldData = new AnnotationFieldData();
                fieldData.setFieldName(cls.getName());
                fieldData.setFieldDescription(title);
                fieldData.setMethod(false);
                return fieldData;
            }
        }
    }

    public FieldData getFieldDataIncludeAnnotation(Field field) {
        JExcelCell fa = field.getAnnotation(JExcelCell.class);
        String title;
        if (fa == null || (title = fa.title().trim()).isEmpty()) {
            FieldData fieldData = new FieldData();
            fieldData.setFieldName(field.getName());
            fieldData.setFieldDescription(field.getName());
            return fieldData;
        } else {
            AnnotationFieldData fieldData = new AnnotationFieldData();
            fieldData.setFieldName(field.getName());
            fieldData.setFieldDescription(title);
            fieldData.setMethod(false);
            return fieldData;
        }
    }

    public FieldData getFieldDataIncludeAnnotation(Method method) {
        JExcelCell ma = method.getAnnotation(JExcelCell.class);
        String title;
        if (ma == null || (title = ma.title().trim()).isEmpty()) {
            FieldData fieldData = new FieldData();
            fieldData.setFieldName(method.getName());
            fieldData.setFieldDescription(method.getName());
            return fieldData;
        } else {
            AnnotationFieldData fieldData = new AnnotationFieldData();
            fieldData.setFieldName(method.getName());
            fieldData.setFieldDescription(title);
            fieldData.setMethod(true);
            if (method.getParameterTypes().length > 0) {
                throw new RuntimeException("Only support no args method");
            }
            return fieldData;
        }
    }

    public boolean isAnnotationOnly() {
        return annotationOnly;
    }

    public void setAnnotationOnly(boolean annotationOnly) {
        this.annotationOnly = annotationOnly;
    }

    public static class AnnotationFieldData extends FieldData {
        private boolean isMethod;

        public boolean isMethod() {
            return isMethod;
        }

        public void setMethod(boolean method) {
            isMethod = method;
        }
    }
}
