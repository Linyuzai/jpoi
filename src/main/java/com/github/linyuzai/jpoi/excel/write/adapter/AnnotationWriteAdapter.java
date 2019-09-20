package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.write.annotation.JExcelWriteCell;
import com.github.linyuzai.jpoi.excel.write.annotation.JExcelWriteSheet;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public abstract class AnnotationWriteAdapter extends ClassWriteAdapter {

    public WriteField getWriteFieldIncludeAnnotation(Class<?> cls) {
        if (cls == null) {
            return null;
        } else {
            JExcelWriteSheet ca = cls.getAnnotation(JExcelWriteSheet.class);
            if (ca == null) {
                return null;
            }
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(cls.getName());
            writeField.setMethod(false);
            writeField.setAnnotationOnly(ca.annotationOnly());
            String title;
            if ((title = ca.name().trim()).isEmpty()) {
                writeField.setFieldDescription(null);
            } else {
                writeField.setFieldDescription(title);
            }
            return writeField;
        }
    }

    public WriteField getWriteFieldIncludeAnnotation(Field field) {
        JExcelWriteCell fa = field.getAnnotation(JExcelWriteCell.class);
        /*if (fa == null) {
            WriteField writeField = new WriteField();
            writeField.setFieldName(field.getName());
            writeField.setFieldDescription(field.getName());
            writeField.setAutoSize(true);
            writeField.setWidth(0);
            writeField.setOrder(Integer.MAX_VALUE);
            return writeField;
        } else {
            String title = fa.title().trim();
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(field.getName());
            writeField.setFieldDescription(title.isEmpty() ? field.getName() : title);
            writeField.setAutoSize(fa.autoSize());
            writeField.setWidth(fa.width());
            writeField.setMethod(false);
            writeField.setOrder(fa.order());
            reuseValueConverter(writeField, fa.valueConverter());
            return writeField;
        }*/
        WriteField writeField = getWriteFieldIncludeAnnotation(field.getName(), fa);
        if (writeField instanceof AnnotationWriteField) {
            ((AnnotationWriteField) writeField).setMethod(false);
        }
        return writeField;
    }

    public WriteField getWriteFieldIncludeAnnotation(Method method) {
        JExcelWriteCell ma = method.getAnnotation(JExcelWriteCell.class);
        /*if (ma == null) {
            WriteField writeField = new WriteField();
            writeField.setFieldName(method.getName());
            writeField.setFieldDescription(method.getName());
            writeField.setAutoSize(true);
            writeField.setWidth(0);
            writeField.setOrder(Integer.MAX_VALUE);
            return writeField;
        } else {
            String title = ma.title().trim();
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(method.getName());
            writeField.setFieldDescription(title.isEmpty() ? method.getName() : title);
            writeField.setAutoSize(ma.autoSize());
            writeField.setWidth(ma.width());
            writeField.setMethod(true);
            writeField.setOrder(ma.order());
            if (method.getParameterTypes().length > 0) {
                throw new RuntimeException("Only support no args method");
            }
            reuseValueConverter(writeField, ma.valueConverter());
            return writeField;
        }*/
        WriteField writeField = getWriteFieldIncludeAnnotation(method.getName(), ma);
        if (writeField instanceof AnnotationWriteField) {
            ((AnnotationWriteField) writeField).setMethod(true);
            if (method.getParameterTypes().length > 0) {
                throw new RuntimeException("Only support no args method");
            }
        }
        return writeField;
    }

    private WriteField getWriteFieldIncludeAnnotation(String name, JExcelWriteCell annotation) {
        if (annotation == null) {
            WriteField writeField = new WriteField();
            writeField.setFieldName(name);
            writeField.setFieldDescription(name);
            writeField.setAutoSize(true);
            writeField.setWidth(0);
            writeField.setOrder(Integer.MAX_VALUE);
            return writeField;
        } else {
            String title = annotation.title().trim();
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(name);
            writeField.setFieldDescription(title.isEmpty() ? name : title);
            writeField.setAutoSize(annotation.autoSize());
            writeField.setWidth(annotation.width());
            writeField.setOrder(annotation.order());
            writeField.setValueConverter(ValueConverter.getWithCache(annotation.valueConverter()));
            //reuseValueConverter(writeField, annotation.valueConverter());
            return writeField;
        }
    }

    private void reuseValueConverter(AnnotationWriteField writeField, Class<? extends ValueConverter> cls) {
        if (!cls.isInterface()) {
            try {
                if (cls == ErrorValueConverter.class) {
                    writeField.setValueConverter(ErrorValueConverter.getInstance());
                } else if (cls == FormulaValueConverter.class) {
                    writeField.setValueConverter(FormulaValueConverter.getInstance());
                } else if (cls == NullValueConverter.class) {
                    writeField.setValueConverter(NullValueConverter.getInstance());
                } else if (cls == WriteObjectValueConverter.class) {
                    writeField.setValueConverter(WriteObjectValueConverter.getInstance());
                } else if (cls == PictureValueConverter.class) {
                    writeField.setValueConverter(PictureValueConverter.getInstance());
                } else if (cls == PoiValueConverter.class) {
                    writeField.setValueConverter(PoiValueConverter.getInstance());
                } else if (cls == WriteSupportValueConverter.class) {
                    writeField.setValueConverter(WriteSupportValueConverter.getInstance());
                } else if (ValueConverter.cache.containsKey(cls.getName())) {
                    writeField.setValueConverter(ValueConverter.cache.get(cls.getName()));
                } else {
                    ValueConverter valueConverter = cls.newInstance();
                    ValueConverter.cache.put(cls.getName(), valueConverter);
                    writeField.setValueConverter(valueConverter);
                }
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    public static class AnnotationWriteField extends WriteField {

        private boolean isMethod;
        private boolean annotationOnly;
        private ValueConverter valueConverter;

        public boolean isMethod() {
            return isMethod;
        }

        public void setMethod(boolean method) {
            isMethod = method;
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
