package com.github.linyuzai.jpoi.excel.write.annotation;

import org.apache.poi.ss.usermodel.BorderStyle;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellBorder {

    BorderStyle top() default BorderStyle.NONE;

    short topColor() default 0;

    BorderStyle right() default BorderStyle.NONE;

    short rightColor() default 0;

    BorderStyle bottom() default BorderStyle.NONE;

    short bottomColor() default 0;

    BorderStyle left() default BorderStyle.NONE;

    short leftColor() default 0;
}
