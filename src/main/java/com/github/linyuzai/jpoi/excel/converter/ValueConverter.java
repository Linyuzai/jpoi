package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.order.Ordered;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public interface ValueConverter extends Ordered {

    Map<String, ValueConverter> cache = new ConcurrentHashMap<>();

    static ValueConverter getWithCache(Class<? extends ValueConverter> cls) {
        if (cls.isInterface()) {
            return null;
        }
        if (cls == CommentValueConverter.class) {
            return CommentValueConverter.getInstance();
        } else if (cls == ErrorValueConverter.class) {
            return ErrorValueConverter.getInstance();
        } else if (cls == FilePictureValueConverter.class) {
            return FilePictureValueConverter.getInstance();
        } else if (cls == FormulaValueConverter.class) {
            return FormulaValueConverter.getInstance();
        } else if (cls == NullValueConverter.class) {
            return NullValueConverter.getInstance();
        } else if (cls == PictureValueConverter.class) {
            return PictureValueConverter.getInstance();
        } else if (cls == PoiValueConverter.class) {
            return PoiValueConverter.getInstance();
        } else if (cls == ReadObjectValueConverter.class) {
            return WriteObjectValueConverter.getInstance();
        } else if (cls == ReadSupportValueConverter.class) {
            return WriteSupportValueConverter.getInstance();
        } else if (cls == WriteObjectValueConverter.class) {
            return WriteObjectValueConverter.getInstance();
        } else if (cls == WriteSupportValueConverter.class) {
            return WriteSupportValueConverter.getInstance();
        } else if (ValueConverter.cache.containsKey(cls.getName())) {
            return ValueConverter.cache.get(cls.getName());
        } else {
            try {
                ValueConverter valueConverter = cls.newInstance();
                ValueConverter.cache.put(cls.getName(), valueConverter);
                return valueConverter;
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
                return null;
            }
        }
    }

    boolean supportValue(int sheet, int row, int cell, Object value);

    Object convertValue(int sheet, int row, int cell, Object value);
}
