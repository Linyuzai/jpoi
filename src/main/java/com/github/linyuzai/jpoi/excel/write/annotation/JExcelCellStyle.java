package com.github.linyuzai.jpoi.excel.write.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

//@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellStyle {

    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.CENTER;

    JExcelCellBorder border() default @JExcelCellBorder;

    short dataFormat() default -1;

    JExcelCellFill fill() default @JExcelCellFill;

    JExcelCellFont font() default @JExcelCellFont;

    boolean hidden() default false;

    short indention() default -1;

    boolean locked() default false;

    boolean quotePrefixed() default false;

    short rotation() default 0;

    boolean shrinkToFit() default false;

    boolean wrapText() default false;
}
