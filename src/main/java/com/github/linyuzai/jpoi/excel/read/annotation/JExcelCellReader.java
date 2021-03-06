package com.github.linyuzai.jpoi.excel.read.annotation;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.lang.annotation.*;

@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellReader {
    /**
     * @return 每一列的标题
     */
    String title() default "";

    /**
     * @return value转换器
     */
    Class<? extends ValueConverter> valueConverter() default ValueConverter.class;

    @Deprecated
    int index() default -1;

    /**
     * @return 作为对应字段的注释
     */
    String commentOfField() default "";

    /**
     * @return 作为对应index列的注释
     */
    int commentOfIndex() default -1;

    /**
     * @return 作为对应字段的图片
     */
    String pictureOfFiled() default "";

    /**
     * @return 作为对应index的图片
     */
    int pictureOfIndex() default -1;
}
