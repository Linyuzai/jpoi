package com.github.linyuzai.jpoi.excel.write.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelRowStyle {

    short height() default 0;

    float heightInPoints() default 0f;

    boolean zeroHeight() default false;

    JExcelCellStyle cellStyle() default @JExcelCellStyle;
}