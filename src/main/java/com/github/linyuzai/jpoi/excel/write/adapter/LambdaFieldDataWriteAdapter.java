package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.common.SerializedLambda;
import com.github.linyuzai.jpoi.exception.JPoiException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public class LambdaFieldDataWriteAdapter extends ListDataWriteAdapter {

    @SafeVarargs
    public final <T> void addListData(ListData listData, Function<T, ?>... lambdaMethods) throws Throwable {
        int sheet = getListDataList().size();
        getListDataList().add(sheet, listData);
        List<WriteField> writeFieldList = new ArrayList<>();
        for (Function<T, ?> lambdaMethod : lambdaMethods) {
            SerializedLambda serializedLambda = SerializedLambda.resolve(lambdaMethod);
            String methodName = serializedLambda.getImplMethodName();
            Method method;
            Class<?> cls = serializedLambda.getImplClass();
            try {
                method = cls.getMethod(methodName);
            } catch (NoSuchMethodException e) {
                throw new JPoiException(e);
            }
            WriteField writeField = getWriteFieldIncludeAnnotation(method);
            if (writeField instanceof AnnotationWriteField) {
                writeFieldList.add(writeField);
            } else {
                String fieldNameFromMethodName = methodName;
                if (methodName.startsWith("get")) {
                    fieldNameFromMethodName = methodName.substring(3);
                    if (!Character.isLowerCase(fieldNameFromMethodName.charAt(0))) {
                        fieldNameFromMethodName = Character.toLowerCase(fieldNameFromMethodName.charAt(0)) +
                                fieldNameFromMethodName.substring(1);
                    }
                }
                Field fieldFromMethodName;
                try {
                    fieldFromMethodName = cls.getField(fieldNameFromMethodName);
                } catch (NoSuchFieldException ignore) {
                    continue;
                }
                writeFieldList.add(getWriteFieldIncludeAnnotation(fieldFromMethodName));
            }
        }
        getWriteFieldList().add(sheet, writeFieldList);
    }
}
