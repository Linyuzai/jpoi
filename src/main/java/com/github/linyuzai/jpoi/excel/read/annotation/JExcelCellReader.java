package com.github.linyuzai.jpoi.excel.read.annotation;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.lang.annotation.*;

@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellReader {

    String title() default "";

    Class<? extends ValueConverter> valueConverter() default ValueConverter.class;

    @Deprecated
    int index() default -1;

    String commentOfField() default "";

    int commentOfIndex() default -1;

    String pictureOfFiled() default "";

    int pictureOfIndex() default -1;
}
