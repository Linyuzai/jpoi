package com.github.linyuzai.jpoi.excel.read.annotation;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.lang.annotation.*;

@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelReadCell {

    String title() default "";

    Class<? extends ValueConverter> valueConverter() default ValueConverter.class;

    int index() default -1;
}
