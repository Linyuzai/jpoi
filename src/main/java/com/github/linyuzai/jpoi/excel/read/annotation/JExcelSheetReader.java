package com.github.linyuzai.jpoi.excel.read.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelSheetReader {

    String name() default "";

    @Deprecated
    boolean annotationOnly() default true;

    boolean toMap() default false;
}
