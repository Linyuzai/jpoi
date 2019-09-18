package com.github.linyuzai.jpoi.excel.write.annotation;

import com.github.linyuzai.jpoi.excel.write.converter.ValueConverter;

import java.lang.annotation.*;

@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelWriteCell {

    String title() default "";

    Class<? extends ValueConverter> valueConverter() default ValueConverter.class;

    boolean autoSize() default true;

    int width() default 0;

    int order() default Integer.MAX_VALUE;
}