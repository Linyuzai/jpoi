package com.github.linyuzai.jpoi.excel.adapter;

import com.github.linyuzai.jpoi.excel.annotation.JExcelCell;
import com.github.linyuzai.jpoi.excel.annotation.JExcelSheet;
import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

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
        if (fa == null) {
            FieldData fieldData = new FieldData();
            fieldData.setFieldName(field.getName());
            fieldData.setFieldDescription(field.getName());
            return fieldData;
        } else {
            String title = fa.title().trim();
            AnnotationFieldData fieldData = new AnnotationFieldData();
            fieldData.setFieldName(field.getName());
            fieldData.setFieldDescription(title.isEmpty() ? field.getName() : title);
            fieldData.setMethod(false);
            reuseValueConverter(fieldData, fa.valueConverter());
            return fieldData;
        }
    }

    public FieldData getFieldDataIncludeAnnotation(Method method) {
        JExcelCell ma = method.getAnnotation(JExcelCell.class);
        if (ma == null) {
            FieldData fieldData = new FieldData();
            fieldData.setFieldName(method.getName());
            fieldData.setFieldDescription(method.getName());
            return fieldData;
        } else {
            String title = ma.title().trim();
            AnnotationFieldData fieldData = new AnnotationFieldData();
            fieldData.setFieldName(method.getName());
            fieldData.setFieldDescription(title.isEmpty() ? method.getName() : title);
            fieldData.setMethod(true);
            if (method.getParameterTypes().length > 0) {
                throw new RuntimeException("Only support no args method");
            }
            reuseValueConverter(fieldData, ma.valueConverter());
            return fieldData;
        }
    }

    private void reuseValueConverter(AnnotationFieldData fieldData, Class<? extends ValueConverter> cls) {
        if (!cls.isInterface()) {
            try {
                if (AnnotationFieldData.valueConverterMap.containsKey(cls.getName())) {
                    fieldData.setValueConverter(AnnotationFieldData.valueConverterMap.get(cls.getName()));
                } else {
                    ValueConverter valueConverter = cls.newInstance();
                    AnnotationFieldData.valueConverterMap.put(cls.getName(), valueConverter);
                    fieldData.setValueConverter(valueConverter);
                }
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    public boolean isAnnotationOnly() {
        return annotationOnly;
    }

    public void setAnnotationOnly(boolean annotationOnly) {
        this.annotationOnly = annotationOnly;
    }

    public static class AnnotationFieldData extends FieldData {

        private static Map<String, ValueConverter> valueConverterMap = new ConcurrentHashMap<>();

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
