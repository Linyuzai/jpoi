package com.github.linyuzai.jpoi.excel.write.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelSheetWriter {

    String name() default "";

    boolean annotationOnly() default true;

    JExcelRowStyle style() default @JExcelRowStyle;
}
