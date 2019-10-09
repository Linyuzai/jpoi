package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.write.annotation.*;
import com.github.linyuzai.jpoi.excel.write.style.*;
import org.apache.poi.ss.usermodel.BuiltinFormats;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public abstract class AnnotationWriteAdapter extends ClassWriteAdapter {

    public WriteField getWriteFieldIncludeAnnotation(Class<?> cls) {
        if (cls == null) {
            return null;
        } else {
            JExcelSheetWriter ca = cls.getAnnotation(JExcelSheetWriter.class);
            if (ca == null) {
                return null;
            }
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(cls.getName());
            writeField.setMethod(false);
            writeField.setAnnotationOnly(ca.annotationOnly());
            writeField.setRowStyle(getRowStyle(ca.style()));
            String title;
            if ((title = ca.name().trim()).isEmpty()) {
                writeField.setFieldDescription(null);
            } else {
                writeField.setFieldDescription(title);
            }
            return writeField;
        }
    }

    public WriteField getWriteFieldIncludeAnnotation(Field field) {
        JExcelCellWriter fa = field.getAnnotation(JExcelCellWriter.class);
        /*if (fa == null) {
            WriteField writeField = new WriteField();
            writeField.setFieldName(field.getName());
            writeField.setFieldDescription(field.getName());
            writeField.setAutoSize(true);
            writeField.setWidth(0);
            writeField.setOrder(Integer.MAX_VALUE);
            return writeField;
        } else {
            String title = fa.title().trim();
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(field.getName());
            writeField.setFieldDescription(title.isEmpty() ? field.getName() : title);
            writeField.setAutoSize(fa.autoSize());
            writeField.setWidth(fa.width());
            writeField.setMethod(false);
            writeField.setOrder(fa.order());
            reuseValueConverter(writeField, fa.valueConverter());
            return writeField;
        }*/
        WriteField writeField = getWriteFieldIncludeAnnotation(field.getName(), fa);
        if (writeField instanceof AnnotationWriteField) {
            ((AnnotationWriteField) writeField).setMethod(false);
        }
        return writeField;
    }

    public WriteField getWriteFieldIncludeAnnotation(Method method) {
        JExcelCellWriter ma = method.getAnnotation(JExcelCellWriter.class);
        /*if (ma == null) {
            WriteField writeField = new WriteField();
            writeField.setFieldName(method.getName());
            writeField.setFieldDescription(method.getName());
            writeField.setAutoSize(true);
            writeField.setWidth(0);
            writeField.setOrder(Integer.MAX_VALUE);
            return writeField;
        } else {
            String title = ma.title().trim();
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(method.getName());
            writeField.setFieldDescription(title.isEmpty() ? method.getName() : title);
            writeField.setAutoSize(ma.autoSize());
            writeField.setWidth(ma.width());
            writeField.setMethod(true);
            writeField.setOrder(ma.order());
            if (method.getParameterTypes().length > 0) {
                throw new RuntimeException("Only support no args method");
            }
            reuseValueConverter(writeField, ma.valueConverter());
            return writeField;
        }*/
        WriteField writeField = getWriteFieldIncludeAnnotation(method.getName(), ma);
        if (writeField instanceof AnnotationWriteField) {
            ((AnnotationWriteField) writeField).setMethod(true);
            if (method.getParameterTypes().length > 0) {
                throw new RuntimeException("Only support no args method");
            }
        }
        return writeField;
    }

    private WriteField getWriteFieldIncludeAnnotation(String name, JExcelCellWriter annotation) {
        if (annotation == null) {
            WriteField writeField = new WriteField();
            writeField.setFieldName(name);
            writeField.setFieldDescription(name);
            writeField.setAutoSize(true);
            writeField.setWidth(0);
            writeField.setOrder(Integer.MAX_VALUE);
            return writeField;
        } else {
            String title = annotation.title().trim();
            AnnotationWriteField writeField = new AnnotationWriteField();
            writeField.setFieldName(name);
            writeField.setFieldDescription(title.isEmpty() ? name : title);
            writeField.setAutoSize(annotation.autoSize());
            writeField.setWidth(annotation.width());
            writeField.setOrder(annotation.order());
            writeField.setValueConverter(ValueConverter.getWithCache(annotation.valueConverter()));
            String commentField = annotation.commentOfField();
            writeField.setCommentOfField(commentField.isEmpty() ? null : commentField);
            writeField.setCommentOfIndex(annotation.commentOfIndex());
            String pictureField = annotation.pictureOfFiled();
            writeField.setPictureOfField(pictureField.isEmpty() ? null : pictureField);
            writeField.setPictureOfIndex(annotation.pictureOfIndex());
            //reuseValueConverter(writeField, annotation.valueConverter());
            writeField.setCellStyle(getCellStyle(annotation.style()));
            return writeField;
        }
    }

    private JRowStyle getRowStyle(JExcelRowStyle annotation) {
        JRowStyle rowStyle = new JRowStyle();
        rowStyle.setHeight(annotation.height());
        rowStyle.setHeightInPoints(annotation.heightInPoints());
        rowStyle.setZeroHeight(annotation.zeroHeight());
        rowStyle.setCellStyle(getCellStyle(annotation.cellStyle()));
        return rowStyle;
    }

    private JCellStyle getCellStyle(JExcelCellStyle annotation) {

        JCellStyle cellStyle = new JCellStyle();

        cellStyle.setHidden(annotation.hidden());
        cellStyle.setHorizontalAlignment(annotation.horizontalAlignment());
        cellStyle.setIndention(annotation.indention());
        cellStyle.setLocked(annotation.locked());
        cellStyle.setQuotePrefixed(annotation.quotePrefixed());
        cellStyle.setRotation(annotation.rotation());
        cellStyle.setShrinkToFit(annotation.shrinkToFit());
        cellStyle.setVerticalAlignment(annotation.verticalAlignment());
        cellStyle.setWrapText(annotation.wrapText());

        short dataFormat = annotation.dataFormat();
        if (dataFormat < 0) {
            dataFormat = (short) BuiltinFormats.getBuiltinFormat(annotation.dataFormatString());
            if (dataFormat < 0) {
                dataFormat = 0;
            }
        }
        cellStyle.setDataFormat(dataFormat);

        JExcelCellBorder border = annotation.border();
        JCellBorder cellBorder = new JCellBorder();
        cellBorder.setBottom(border.bottom());
        cellBorder.setBottomColor(border.bottomColor());
        cellBorder.setLeft(border.left());
        cellBorder.setLeftColor(border.leftColor());
        cellBorder.setRight(border.right());
        cellBorder.setRightColor(border.rightColor());
        cellBorder.setTop(border.top());
        cellBorder.setTopColor(border.topColor());
        cellStyle.setBorder(cellBorder);

        JExcelCellFill fill = annotation.fill();
        JCellFill cellFill = new JCellFill();
        cellFill.setPattern(fill.pattern());
        cellFill.setBackgroundColor(fill.backgroundColor());
        cellFill.setForegroundColor(fill.foregroundColor());
        cellStyle.setFill(cellFill);

        JExcelCellFont font = annotation.font();
        JCellFont cellFont = new JCellFont();
        cellFont.setBold(font.bold());
        cellFont.setCharSet(font.charSet());
        cellFont.setColor(font.color());
        cellFont.setFontHeight(font.fontHeight());
        cellFont.setFontHeightInPoints(font.fontHeightInPoints());
        cellFont.setFontName(font.fontName());
        cellFont.setItalic(font.italic());
        cellFont.setStrikeout(font.strikeout());
        cellFont.setTypeOffset(font.typeOffset());
        cellFont.setUnderline(font.underline());
        cellStyle.setFont(cellFont);

        return cellStyle;
    }

    public static class AnnotationWriteField extends WriteField {

        private boolean isMethod;
        private boolean annotationOnly;
        private ValueConverter valueConverter;
        private String commentOfField;
        private int commentOfIndex;
        private String pictureOfField;
        private int pictureOfIndex;
        private JRowStyle rowStyle;
        private JCellStyle cellStyle;
        private List<WriteField> combinationFields = new ArrayList<>();

        public boolean isMethod() {
            return isMethod;
        }

        public void setMethod(boolean method) {
            isMethod = method;
        }

        public boolean isAnnotationOnly() {
            return annotationOnly;
        }

        public void setAnnotationOnly(boolean annotationOnly) {
            this.annotationOnly = annotationOnly;
        }

        public ValueConverter getValueConverter() {
            return valueConverter;
        }

        public void setValueConverter(ValueConverter valueConverter) {
            this.valueConverter = valueConverter;
        }

        public String getCommentOfField() {
            return commentOfField;
        }

        public void setCommentOfField(String commentOfField) {
            this.commentOfField = commentOfField;
        }

        public int getCommentOfIndex() {
            return commentOfIndex;
        }

        public void setCommentOfIndex(int commentOfIndex) {
            this.commentOfIndex = commentOfIndex;
        }

        public String getPictureOfField() {
            return pictureOfField;
        }

        public void setPictureOfField(String pictureOfField) {
            this.pictureOfField = pictureOfField;
        }

        public int getPictureOfIndex() {
            return pictureOfIndex;
        }

        public void setPictureOfIndex(int pictureOfIndex) {
            this.pictureOfIndex = pictureOfIndex;
        }

        public JRowStyle getRowStyle() {
            return rowStyle;
        }

        public void setRowStyle(JRowStyle rowStyle) {
            this.rowStyle = rowStyle;
        }

        public JCellStyle getCellStyle() {
            return cellStyle;
        }

        public void setCellStyle(JCellStyle cellStyle) {
            this.cellStyle = cellStyle;
        }

        public List<WriteField> getCombinationFields() {
            return combinationFields;
        }

        public void setCombinationFields(List<WriteField> combinationFields) {
            this.combinationFields = combinationFields;
        }

        public boolean isComment() {
            return commentOfIndex >= 0 || (commentOfField != null && !commentOfField.isEmpty());
        }

        public boolean isPicture() {
            return pictureOfIndex >= 0 || (pictureOfField != null && !pictureOfField.isEmpty());
        }
    }
}
