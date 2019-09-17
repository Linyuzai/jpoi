package com.github.linyuzai.jpoi.excel.adapter.write;

import com.github.linyuzai.jpoi.excel.annotation.JExcelCell;
import com.github.linyuzai.jpoi.excel.annotation.JExcelSheet;
import com.github.linyuzai.jpoi.excel.converter.*;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public abstract class AnnotationWriteAdapter extends ClassWriteAdapter {

    private boolean annotationOnly;

    public FieldData getFieldDataIncludeAnnotation(Class<?> cls) {
        if (cls == null) {
            return null;
        } else {
            JExcelSheet ca = cls.getAnnotation(JExcelSheet.class);
            if (ca == null) {
                return null;
            }
            String title;
            annotationOnly = ca.annotationOnly();
            if ((title = ca.name().trim()).isEmpty()) {
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
            fieldData.setAutoSize(true);
            fieldData.setOrder(Integer.MAX_VALUE);
            return fieldData;
        } else {
            String title = fa.title().trim();
            AnnotationFieldData fieldData = new AnnotationFieldData();
            fieldData.setFieldName(field.getName());
            fieldData.setFieldDescription(title.isEmpty() ? field.getName() : title);
            fieldData.setAutoSize(fa.autoSize());
            fieldData.setMethod(false);
            fieldData.setOrder(fa.order());
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
            fieldData.setAutoSize(true);
            fieldData.setOrder(Integer.MAX_VALUE);
            return fieldData;
        } else {
            String title = ma.title().trim();
            AnnotationFieldData fieldData = new AnnotationFieldData();
            fieldData.setFieldName(method.getName());
            fieldData.setFieldDescription(title.isEmpty() ? method.getName() : title);
            fieldData.setAutoSize(ma.autoSize());
            fieldData.setMethod(true);
            fieldData.setOrder(ma.order());
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
                if (cls == NullValueConverter.class) {
                    fieldData.setValueConverter(NullValueConverter.getInstance());
                } else if (cls == ObjectValueConverter.class) {
                    fieldData.setValueConverter(ObjectValueConverter.getInstance());
                } else if (cls == PictureValueConverter.class) {
                    fieldData.setValueConverter(PictureValueConverter.getInstance());
                } else if (cls == PoiValueConverter.class) {
                    fieldData.setValueConverter(PoiValueConverter.getInstance());
                } else if (cls == SupportValueConverter.class) {
                    fieldData.setValueConverter(SupportValueConverter.getInstance());
                } else if (ValueConverter.cache.containsKey(cls.getName())) {
                    fieldData.setValueConverter(ValueConverter.cache.get(cls.getName()));
                } else {
                    ValueConverter valueConverter = cls.newInstance();
                    ValueConverter.cache.put(cls.getName(), valueConverter);
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
