package com.github.linyuzai.jpoi.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

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

    public static List<Field> getFields(Class<?> cls) {
        List<Field> fields = new ArrayList<>();
        while (cls != null) {
            fields.addAll(Arrays.asList(cls.getDeclaredFields()));
            cls = cls.getSuperclass();
        }
        return fields;
    }

    public static Field getField(Class<?> cls, String name) throws NoSuchFieldException {
        Field field = null;
        while (cls != null) {
            try {
                field = cls.getDeclaredField(name);
                break;
            } catch (NoSuchFieldException ignore) {
                cls = cls.getSuperclass();
            }
        }
        if (field == null) {
            throw new NoSuchFieldException(name);
        }
        return field;
    }

    public static List<Method> getMethods(Class<?> cls) {
        List<Method> methods = new ArrayList<>();
        while (cls != null) {
            methods.addAll(Arrays.asList(cls.getDeclaredMethods()));
            cls = cls.getSuperclass();
        }
        return methods;
    }

    public static Method getMethod(Class<?> cls, String name) throws NoSuchMethodException {
        Method method = null;
        while (cls != null) {
            try {
                method = cls.getDeclaredMethod(name);
                break;
            } catch (NoSuchMethodException ignore) {
                cls = cls.getSuperclass();
            }
        }
        if (method == null) {
            throw new NoSuchMethodException(name);
        }
        return method;
    }
}
