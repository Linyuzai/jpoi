# jpoi

### 依赖
```java
implementation 'com.github.linyuzai:jpoi:0.1.0'
```

### 基本用法
```java
JExcel.xlsx().setWriteAdapter(WriteAdapter).write().to(File/OutputStream);//写

Object v = JExcel.xlsx(InputStream).setReadAdapter(ReadAdapter).read().getValue();//读
```
### 支持用法
```java
JExcel.xlsx().data(List...).write().to(File/OutputStream);//写类

List<List<Class>> v = JExcel.xlsx(InputStream).target(Class).read().getValue();//读类

List<List<Map<String, Object>>> v = JExcel.xlsx(InputStream).toMap().read().getValue();//读map

List<List<List<Object>>> v = JExcel.xlsx(InputStream).direct().read().getValue();//读list
```
### 全部用法
```java
JExcel
      .xls()//HSSFWorkbook
      .xlsx()//XSSFWorkbook
      .sxlsx()//SXSSFWorkbook
      .data()//list bean
      .setWriteAdapter()//自定义WriteAdapter
      .addPoiListener()//poi监听器
      .addValueConverter()//value转换器
      .setValueSetter()//value写入poi的支持类
      .write()//执行写入
      .to();//输出

JExcel
      .xls(InputStream)//HSSFWorkbook
      .xlsx(InputStream)//XSSFWorkbook
      .sxlsx(InputStream)//自定义Sax写法，暂不完善，无法使用
      .target()//转成bean
      .toMap()//转成map
      .direct()//转成list
      .setReadAdapter()//自定义ReadAdapter
      .addPoiListener()//poi监听器
      .addValueConverter()//value转换器
      .setValueGetter()//poi读取value的支持类
      .read()//执行读取
      .getValue();//获得值
```
### 注解加类写
```java
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
```
```java
@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface JExcelCellWriter {
    /**
     * @return 每一列的标题
     */
    String title() default "";

    /**
     * @return value转换器
     */
    Class<? extends ValueConverter> valueConverter() default ValueConverter.class;

    /**
     * @return 是否自适应列宽
     */
    boolean autoSize() default true;

    /**
     * @return 指定宽度
     */
    int width() default 0;

    /**
     * @return 排序
     */
    int order() default Integer.MAX_VALUE;

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

    @Deprecated
    String standbyFor() default "";

    /**
     * @return 样式
     */
    JExcelCellStyle style() default @JExcelCellStyle;
}
```
```java
JExcel.xlsx().data(List...).write().to(File/OutputStream);//写类
```
### 注解加类读
```java
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
```
```java
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
```
```java
List<List<Class>> v = JExcel.xlsx(InputStream).target(Class).read().getValue();//读类
```
### 注意事项
- `commentOfField` `commentOfIndex` `pictureOfFiled` `pictureOfIndex`只用于支持一个单元格内的多个数据（内容，注释，图片）对应bean的多个字段，如果只想写入一个值或读取一个值，请直接设置ValueConverter
