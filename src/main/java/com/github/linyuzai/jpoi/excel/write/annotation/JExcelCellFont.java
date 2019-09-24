package com.github.linyuzai.jpoi.excel.write.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellFont {

    boolean bold() default false;

    int charSet() default 0;

    short color() default 0;

    short fontHeight() default 0;

    short fontHeightInPoints() default 0;

    String fontName() default "";

    boolean italic() default false;

    boolean strikeout() default false;

    short typeOffset() default 0;

    byte underline() default 0;
}
