package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.excel.converter.WriteCommentValueConverter;
import com.github.linyuzai.jpoi.excel.converter.WritePictureValueConverter;
import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.listener.ExcelListener;
import com.github.linyuzai.jpoi.excel.value.combination.ListCombinationValue;
import com.github.linyuzai.jpoi.excel.write.style.JCellStyle;
import com.github.linyuzai.jpoi.excel.write.style.JRowStyle;
import com.github.linyuzai.jpoi.exception.JPoiException;
import com.github.linyuzai.jpoi.util.ClassUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

public class ListDataWriteAdapter extends AnnotationWriteAdapter implements ExcelListener {

    private List<ListData> listDataList = new ArrayList<>();
    private List<List<WriteField>> writeFieldList = new ArrayList<>();

    private Workbook workbook;

    public ListDataWriteAdapter() {
    }

    public ListDataWriteAdapter(List<ListData> listDataList) {
        if (listDataList == null) {
            throw new JPoiException("ListDataList is null");
        }
        for (ListData listData : listDataList) {
            addListData(listData);
        }
    }

    public void addListData(ListData listData) {
        addListData(listData, null);
    }

    public void addListData(ListData listData, Class<?> cls) {
        if (listData == null) {
            throw new JPoiException("ListData is null");
        }
        int sheet = listDataList.size();
        listDataList.add(sheet, listData);
        cellFields(sheet, listData.getDataList(), cls);
        combineFields(sheet);
    }

    public List<ListData> getListDataList() {
        return listDataList;
    }

    public List<List<WriteField>> getWriteFieldList() {
        return writeFieldList;
    }

    public WriteField getWriteField(int s, int c) {
        return writeFieldList.get(s).get(c);
    }

    public void addWriteFieldOrdered(int sheet, WriteField writeField) {
        final List<WriteField> writeFields = writeFieldList.get(sheet);
        writeField.setOrder(writeFields.size());
        writeFields.add(writeField);
    }

    public void cellFields(int sheet, List<?> dataList, Class<?> cls) {
        WriteField writeField = getWriteFieldIncludeAnnotation(cls);
        if (writeField instanceof AnnotationWriteField) {
            /*StringBuilder sheetName = new StringBuilder(writeField.getFieldDescription());
            for (ListData listData : listDataList) {
                if (sheetName.toString().equals(listData.getSheetName())) {
                    sheetName.append(sheet);
                    break;
                }
            }*/
            listDataList.get(sheet).setAnnotationOnly(((AnnotationWriteField) writeField).isAnnotationOnly());
            listDataList.get(sheet).setRowStyle(((AnnotationWriteField) writeField).getRowStyle());
            if (writeField.getFieldDescription() != null) {
                listDataList.get(sheet).setSheetName(writeField.getFieldDescription());
            }
        }
        if (cls != null && !cls.isInterface() && !cls.isArray() && !cls.isPrimitive() && !ClassUtils.isWrapClass(cls) && cls != String.class) {
            writeFieldList.add(sheet, getWriteFieldList(sheet, cls));
        } else {
            if (dataList.size() > 0) {
                //writeFieldList.add(sheet, getWriteFieldList(dataList.get(0).getClass()));
                cellFields(sheet, dataList, dataList.get(0).getClass());
            } else {
                writeFieldList.add(sheet, Collections.emptyList());
            }
        }
        for (List<WriteField> fd : writeFieldList) {
            fd.sort(Comparator.comparingInt(WriteField::getOrder));
        }
    }

    public void combineFields(int sheet) {
        List<WriteField> writeFields = writeFieldList.get(sheet);
        Iterator<WriteField> iterator = writeFields.iterator();
        while (iterator.hasNext()) {
            WriteField writeField = iterator.next();
            if (writeField instanceof AnnotationWriteField) {
                boolean isComment = ((AnnotationWriteField) writeField).isComment();
                boolean isPicture = ((AnnotationWriteField) writeField).isPicture();
                String commentField = ((AnnotationWriteField) writeField).getCommentOfField();
                int commentIndex = ((AnnotationWriteField) writeField).getCommentOfIndex();
                ValueConverter valueConverter = ((AnnotationWriteField) writeField).getValueConverter();
                if (isComment) {
                    if (valueConverter == null) {
                        ((AnnotationWriteField) writeField).setValueConverter(WriteCommentValueConverter.getInstance());
                    }
                }
                String pictureField = ((AnnotationWriteField) writeField).getPictureOfField();
                int pictureIndex = ((AnnotationWriteField) writeField).getPictureOfIndex();
                if (isPicture) {
                    if (valueConverter == null) {
                        ((AnnotationWriteField) writeField).setValueConverter(WritePictureValueConverter.getInstance());
                    }
                }
                if (isComment && isPicture) {
                    throw new JPoiException("Can't set comments and pictures at the same time on the field '" + writeField.getFieldName() + "'");
                } else if (isComment || isPicture) {
                    int index = commentIndex >= 0 ? commentIndex : pictureIndex;
                    if (index >= 0) {
                        if (index < writeFields.size()) {
                            WriteField target = writeFields.get(index);
                            if (target instanceof AnnotationWriteField) {
                                ((AnnotationWriteField) target).getCombinationFields().add(writeField);
                            }
                        }
                    } else {
                        String field = (commentField != null && !commentField.isEmpty()) ? commentField : pictureField;
                        for (WriteField rf : writeFields) {
                            if (rf instanceof AnnotationWriteField && field.equals(rf.getFieldName())) {
                                ((AnnotationWriteField) rf).getCombinationFields().add(writeField);
                                break;
                            }
                        }
                    }
                    iterator.remove();
                }
            }
        }
    }

    public List<WriteField> getWriteFieldList(int sheet, Class<?> cls) {
        /*if (Map.class.isAssignableFrom(cls)) {

        }*/
        if (illegalWriteFieldClass(cls)) {
            return Collections.emptyList();
        }
        //Field[] fields = cls.getDeclaredFields();
        List<Field> fields = ClassUtils.getFields(cls);
        List<WriteField> writeFieldList = new ArrayList<>();
        for (Field field : fields) {
            WriteField writeField = getWriteFieldIncludeAnnotation(field);
            if (listDataList.get(sheet).isAnnotationOnly()) {
                if (writeField instanceof AnnotationWriteField) {
                    writeFieldList.add(writeField);
                }
            } else {
                writeFieldList.add(writeField);
            }
        }
        //Method[] methods = cls.getDeclaredMethods();
        List<Method> methods = ClassUtils.getMethods(cls);
        for (Method method : methods) {
            WriteField writeField = getWriteFieldIncludeAnnotation(method);
            if (writeField instanceof AnnotationWriteField) {
                writeFieldList.add(writeField);
            }
        }
        return writeFieldList;
    }

    public boolean illegalWriteFieldClass(Class<?> cls) {
        return cls == null || cls.isInterface() || cls.isArray() || cls.isPrimitive() || ClassUtils.isWrapClass(cls) || cls == String.class;
    }

    @Override
    public int getHeaderRowCount(int sheet) {
        return 0;
    }

    @Override
    public int getHeaderCellCount(int sheet, int row) {
        return 0;
    }

    @Override
    public int getDataRowCount(int sheet) {
        return listDataList.get(sheet).getDataList().size();
    }

    @Override
    public int getDataCellCount(int sheet, int row) {
        return writeFieldList.get(sheet).size();
    }

    @Override
    public Object getHeaderRowContent(int sheet, int row, int cell, int realRow, int realCell) {
        return null;
    }

    @Override
    public Object getHeaderCellContent(int sheet, int row, int cell, int realRow, int realCell) {
        return null;
    }

    @Override
    public Object getDataCalculateHeader(int sheet, int row, int cell, int realRow, int realCell) {
        Object entity = listDataList.get(sheet).getDataList().get(row);
        return getValueFromEntity(sheet, row, cell, realRow, realCell, writeFieldList.get(sheet).get(cell), entity);
    }

    public Object getValueFromEntity(int sheet, int row, int cell, int realRow, int realCell, WriteField writeField, Object entity) {
        String fieldName = writeField.getFieldName();
        Object val;
        if (writeField instanceof AnnotationWriteField && ((AnnotationWriteField) writeField).isMethod()) {
            try {
                //Method method = entity.getClass().getDeclaredMethod(fieldName);
                Method method = ClassUtils.getMethod(entity.getClass(), fieldName);
                method.setAccessible(true);
                val = method.invoke(entity);
            } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                throw new JPoiException(e);
            }
        } else {
            try {
                //Field field = entity.getClass().getDeclaredField(fieldName);
                Field field = ClassUtils.getField(entity.getClass(), fieldName);
                field.setAccessible(true);
                val = field.get(entity);
            } catch (NoSuchFieldException | IllegalAccessException e) {
                throw new JPoiException(e);
            }
        }
        ValueConverter valueConverter;
        if (writeField instanceof AnnotationWriteField &&
                (valueConverter = ((AnnotationWriteField) writeField).getValueConverter()) != null) {
            val = valueConverter.convertValue(sheet, realRow, realCell, val);
        }
        if (writeField instanceof AnnotationWriteField && ((AnnotationWriteField) writeField).getCombinationFields().size() > 0) {
            ListCombinationValue combinationValue = new ListCombinationValue();
            combinationValue.addValue(val);
            for (WriteField combinationField : ((AnnotationWriteField) writeField).getCombinationFields()) {
                combinationValue.addValue(getValueFromEntity(sheet, row, cell, realRow, realCell, combinationField, entity));
            }
            return combinationValue;
        } else {
            return val;
        }
    }

    @Override
    public String getSheetName(int sheet) {
        String sheetName = listDataList.get(sheet).getSheetName();
        sheetName = sheetName == null ? "Sheet" + sheet : sheetName;
        int eq = 0;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet s = workbook.getSheetAt(i);
            if (sheetName.equals(s.getSheetName())) {
                eq++;
            }
        }
        if (eq == 0) {
            return sheetName;
        } else {
            return sheetName + "(" + eq + ")";
        }
    }

    @Override
    public int getSheetCount() {
        return listDataList.size();
    }

    @Override
    public void onWorkbookStart(Workbook workbook, CreationHelper creationHelper) {
        this.workbook = workbook;
    }

    @Override
    public void onSheetStart(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }
    }

    @Override
    public void onRowStart(int r, int s, Row row, Sheet sheet, Workbook workbook) {
        JRowStyle rs = listDataList.get(s).getRowStyle();
        if (rs == null) {
            return;
        }
        CellStyle cellStyle = workbook.createCellStyle();
        CellStyle old = row.getRowStyle();
        if (old != null) {
            cellStyle.cloneStyleFrom(old);
        }
        fillCellStyle(cellStyle, rs.getCellStyle(), workbook);
        row.setRowStyle(cellStyle);
        row.setZeroHeight(rs.isZeroHeight());
        row.setHeightInPoints(rs.getHeightInPoints());
        row.setHeight(rs.getHeight());
    }

    @Override
    public void onCellStart(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        WriteField writeField = writeFieldList.get(s).get(c);
        if (writeField instanceof AnnotationWriteField) {
            JCellStyle cs = ((AnnotationWriteField) writeField).getCellStyle();
            if (cs == null) {
                return;
            }
            CellStyle cellStyle = workbook.createCellStyle();
            CellStyle old = cell.getCellStyle();
            if (old != null) {
                cellStyle.cloneStyleFrom(old);
            }
            fillCellStyle(cellStyle, cs, workbook);
            cell.setCellStyle(cellStyle);
        }
    }

    @Override
    public void onCellEnd(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        if (cell.getCellType() == CellType.STRING) {
            int length = cell.getStringCellValue().getBytes().length;
            WriteField writeField = writeFieldList.get(s).get(c);
            if (length > 0 && writeField.isAutoSize() && writeField.getWidth() < length) {
                writeField.setWidth(length);
            }
        }
    }

    @Override
    public void onSheetEnd(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        for (int columnNum = 0; columnNum < writeFieldList.get(s).size(); columnNum++) {
            WriteField writeField = writeFieldList.get(s).get(columnNum);
            if (writeField.isAutoSize()) {
                sheet.autoSizeColumn(columnNum);
            }
            if (writeField.getWidth() > 0) {//255*256
                int w = writeField.getWidth() * 256 + 200;
                if (w > 255 * 256) {
                    w = 255 * 256;
                }
                sheet.setColumnWidth(columnNum, w);
            }
        }
    }

    private void fillCellStyle(CellStyle style, JCellStyle source, Workbook workbook) {
        style.setAlignment(source.getHorizontalAlignment());
        style.setBorderBottom(source.getBorder().getBottom());
        style.setBottomBorderColor(source.getBorder().getBottomColor());
        style.setBorderLeft(source.getBorder().getLeft());
        style.setLeftBorderColor(source.getBorder().getLeftColor());
        style.setBorderRight(source.getBorder().getRight());
        style.setRightBorderColor(source.getBorder().getRightColor());
        style.setBorderTop(source.getBorder().getTop());
        style.setTopBorderColor(source.getBorder().getTopColor());
        style.setDataFormat(source.getDataFormat());
        style.setFillPattern(source.getFill().getPattern());
        style.setFillBackgroundColor(source.getFill().getBackgroundColor());
        style.setFillForegroundColor(source.getFill().getForegroundColor());
        style.setHidden(source.isHidden());
        style.setIndention(source.getIndention());
        style.setLocked(source.isLocked());
        style.setQuotePrefixed(source.isQuotePrefixed());
        style.setRotation(source.getRotation());
        style.setShrinkToFit(source.isShrinkToFit());
        style.setWrapText(source.isWrapText());
        style.setVerticalAlignment(source.getVerticalAlignment());
        /*Font font = workbook.getFontAt(style.getFontIndexAsInt());
        if (font == null) {
            font = workbook.createFont();
        }*/
        Font font = workbook.createFont();
        font.setBold(source.getFont().isBold());
        font.setCharSet(source.getFont().getCharSet());
        //font.setCharSet();
        font.setColor(source.getFont().getColor());
        font.setFontHeight(source.getFont().getFontHeight());
        font.setFontHeightInPoints(source.getFont().getFontHeightInPoints());
        font.setFontName(source.getFont().getFontName());
        font.setItalic(source.getFont().isItalic());
        font.setStrikeout(source.getFont().isStrikeout());
        font.setTypeOffset(source.getFont().getTypeOffset());
        font.setUnderline(source.getFont().getUnderline());
        style.setFont(font);
    }

    @Override
    public int getOrder() {
        return Integer.MIN_VALUE;
    }

    public static class ListData {
        private String sheetName;
        private boolean annotationOnly;
        private JRowStyle rowStyle;
        private List<?> dataList;

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public JRowStyle getRowStyle() {
            return rowStyle;
        }

        public void setRowStyle(JRowStyle rowStyle) {
            this.rowStyle = rowStyle;
        }

        public boolean isAnnotationOnly() {
            return annotationOnly;
        }

        public void setAnnotationOnly(boolean annotationOnly) {
            this.annotationOnly = annotationOnly;
        }

        public List<?> getDataList() {
            return dataList;
        }

        public void setDataList(List<?> dataList) {
            this.dataList = dataList;
        }
    }
}
