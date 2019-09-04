package com.github.linyuzai.jpoi.excel.annotation;

import java.lang.annotation.*;

@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCell {

    String title() default "";
}
