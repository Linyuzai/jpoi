package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.support.SupportCache;
import com.github.linyuzai.jpoi.support.SupportOrder;

import java.lang.reflect.Modifier;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public interface ValueConverter extends SupportOrder, SupportCache {

    Map<String, ValueConverter> cache = new ConcurrentHashMap<>();

    static ValueConverter getWithCache(Class<? extends ValueConverter> cls) {
        if (cls == BaseTypeValueConverter.class) {
            return BaseTypeValueConverter.getInstance();
        } else if (cls == DefaultReadValueConverter.class) {
            return DefaultReadValueConverter.getInstance();
        } else if (cls == ErrorValueConverter.class) {
            return ErrorValueConverter.getInstance();
        } else if (cls == FilePictureValueConverter.class) {
            return FilePictureValueConverter.getInstance();
        } else if (cls == NullValueConverter.class) {
            return NullValueConverter.getInstance();
        } else if (cls == PoiValueConverter.class) {
            return PoiValueConverter.getInstance();
        } else if (cls == PostValueConverter.class) {
            return PostValueConverter.getInstance();
        } else if (cls == ReadBase64PictureValueConverter.class) {
            return ReadBase64PictureValueConverter.getInstance();
        } else if (cls == ReadCommentValueConverter.class) {
            return ReadCommentValueConverter.getInstance();
        } else if (cls == ReadDataFormatValueConverter.class) {
            return ReadDataFormatValueConverter.getInstance();
        } else if (cls == ReadDataValueConverter.class) {
            return ReadDataValueConverter.getInstance();
        } else if (cls == ReadFormulaValueConverter.class) {
            return ReadFormulaValueConverter.getInstance();
        } else if (cls == ReadObjectValueConverter.class) {
            return ReadObjectValueConverter.getInstance();
        } else if (cls == ReadPictureValueConverter.class) {
            return ReadPictureValueConverter.getInstance();
        } else if (cls == ReadSupportValueConverter.class) {
            return ReadSupportValueConverter.getInstance();
        } else if (cls == ReadUrlPictureValueConverter.class) {
            return ReadUrlPictureValueConverter.getInstance();
        } else if (cls == StringNumberValueConverter.class) {
            return StringNumberValueConverter.getInstance();
        } else if (cls == StringValueConverter.class) {
            return StringValueConverter.getInstance();
        } else if (cls == WriteCombinationValueConverter.class) {
            return WriteCombinationValueConverter.getInstance();
        } else if (cls == WriteCommentValueConverter.class) {
            return WriteCommentValueConverter.getInstance();
        } else if (cls == WriteFormulaValueConverter.class) {
            return WriteFormulaValueConverter.getInstance();
        } else if (cls == WriteObjectValueConverter.class) {
            return WriteObjectValueConverter.getInstance();
        } else if (cls == WritePictureValueConverter.class) {
            return WritePictureValueConverter.getInstance();
        } else if (cls == WriteSupportValueConverter.class) {
            return WriteSupportValueConverter.getInstance();
        } else if (cls == WriteUrlPictureValueConverter.class) {
            return WriteUrlPictureValueConverter.getInstance();
        } else if (ValueConverter.cache.containsKey(cls.getName())) {
            return ValueConverter.cache.get(cls.getName());
        } else {
            if (cls.isInterface() || Modifier.isAbstract(cls.getModifiers())) {
                return null;
            }
            try {
                ValueConverter valueConverter = cls.newInstance();
                if (valueConverter.isCacheEnabled()) {
                    ValueConverter.cache.put(cls.getName(), valueConverter);
                }
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
