package com.github.linyuzai.jpoi.excel.read.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelReadSheet {

    String name() default "";

    boolean annotationOnly() default false;
}
