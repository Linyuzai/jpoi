package com.github.linyuzai.jpoi.excel.read.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelSheetReader {

    @Deprecated
    String name() default "";

    /**
     * @return 是否只处理添加了注解的字段
     */
    @Deprecated
    boolean annotationOnly() default true;

    /**
     * @return 是否转成map
     */
    boolean toMap() default false;
}
