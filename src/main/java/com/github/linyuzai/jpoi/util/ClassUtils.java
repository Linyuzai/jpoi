package com.github.linyuzai.jpoi.util;

public class ClassUtils {

    public static boolean isWrapClass(Class<?> clz) {
        Class baseType = getBaseType(clz);
        if (baseType == null) {
            return false;
        }
        return baseType.isPrimitive();
    }

    public static Class getBaseType(Class<?> clz) {
        try {
            return (Class) clz.getField("TYPE").get(null);
        } catch (Exception e) {
            return null;
        }
    }
}
