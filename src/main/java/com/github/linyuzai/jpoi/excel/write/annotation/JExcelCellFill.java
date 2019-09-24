package com.github.linyuzai.jpoi.excel.write.annotation;

import org.apache.poi.ss.usermodel.FillPatternType;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellFill {

    FillPatternType pattern() default FillPatternType.NO_FILL;

    short foregroundColor() default 0;

    short backgroundColor() default 0;
}
