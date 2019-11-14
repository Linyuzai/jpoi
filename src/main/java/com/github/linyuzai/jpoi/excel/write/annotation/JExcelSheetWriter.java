package com.github.linyuzai.jpoi.excel.write.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelSheetWriter {
    /**
     * @return sheet名称
     */
    String name() default "";

    /**
     * @return 是否只处理添加了注解的字段
     */
    @Deprecated
    boolean annotationOnly() default true;

    /**
     * @return 样式
     */
    JExcelRowStyle style() default @JExcelRowStyle;
}
