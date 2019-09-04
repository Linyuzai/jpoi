package com.github.linyuzai.jpoi.excel.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelSheet {

    String name() default "";
}
