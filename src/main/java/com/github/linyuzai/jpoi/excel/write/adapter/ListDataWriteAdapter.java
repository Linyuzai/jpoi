package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.excel.write.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.write.listener.PoiWriteListener;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

public class ListDataWriteAdapter extends AnnotationWriteAdapter implements PoiWriteListener {

    private List<ListData> listDataList = new ArrayList<>();
    private List<List<WriteField>> writeFieldList = new ArrayList<>();

    private Workbook workbook;

    public ListDataWriteAdapter() {
    }

    public ListDataWriteAdapter(List<ListData> listDataList) {
        if (listDataList == null) {
            throw new RuntimeException("ListDataList is null");
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
            throw new RuntimeException("ListData is null");
        }
        int sheet = listDataList.size();
        listDataList.add(sheet, listData);
        cellFields(sheet, listData.getDataList(), cls);
    }

    public List<ListData> getListDataList() {
        return listDataList;
    }

    public List<List<WriteField>> getWriteFieldList() {
        return writeFieldList;
    }

    public void cellFields(int sheet, List<?> dataList, Class<?> cls) {
        WriteField writeField = getWriteFieldIncludeAnnotation(cls);
        if (writeField != null) {
            /*StringBuilder sheetName = new StringBuilder(writeField.getFieldDescription());
            for (ListData listData : listDataList) {
                if (sheetName.toString().equals(listData.getSheetName())) {
                    sheetName.append(sheet);
                    break;
                }
            }*/
            listDataList.get(sheet).setSheetName(writeField.getFieldDescription());
        }
        if (cls != null && !cls.isInterface() && !cls.isArray() && !cls.isPrimitive() && !isWrapClass(cls) && cls != String.class) {
            writeFieldList.add(sheet, getWriteFieldList(cls));
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

    public List<WriteField> getWriteFieldList(Class<?> cls) {
        /*if (Map.class.isAssignableFrom(cls)) {

        }*/
        if (illegalWriteFieldClass(cls)) {
            return Collections.emptyList();
        }
        Field[] fields = cls.getDeclaredFields();
        List<WriteField> writeFieldList = new ArrayList<>();
        for (Field field : fields) {
            WriteField writeField = getWriteFieldIncludeAnnotation(field);
            if (isAnnotationOnly()) {
                if (writeField instanceof AnnotationWriteField) {
                    writeFieldList.add(writeField);
                }
            } else {
                writeFieldList.add(writeField);
            }
        }
        Method[] methods = cls.getDeclaredMethods();
        for (Method method : methods) {
            WriteField writeField = getWriteFieldIncludeAnnotation(method);
            if (writeField instanceof AnnotationWriteField) {
                writeFieldList.add(writeField);
            }
        }
        return writeFieldList;
    }

    public boolean illegalWriteFieldClass(Class<?> cls) {
        return cls == null || cls.isInterface() || cls.isArray() || cls.isPrimitive() || isWrapClass(cls) || cls == String.class;
    }

    public static boolean isWrapClass(Class<?> clz) {
        try {
            return ((Class) clz.getField("TYPE").get(null)).isPrimitive();
        } catch (Exception e) {
            return false;
        }
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
        return getValueFromEntity(sheet, row, cell, realRow, realCell, entity);
    }

    public Object getValueFromEntity(int sheet, int row, int cell, int realRow, int realCell, Object entity) {
        WriteField writeField = writeFieldList.get(sheet).get(cell);
        String fieldName = writeField.getFieldName();
        Object val;
        if (writeField instanceof AnnotationWriteField && ((AnnotationWriteField) writeField).isMethod()) {
            try {
                Method method = entity.getClass().getMethod(fieldName);
                method.setAccessible(true);
                val = method.invoke(entity);
            } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                throw new RuntimeException(e);
            }
        } else {
            try {
                Field field = entity.getClass().getDeclaredField(fieldName);
                field.setAccessible(true);
                val = field.get(entity);
            } catch (NoSuchFieldException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }
        ValueConverter valueConverter;
        if (writeField instanceof AnnotationWriteField &&
                (valueConverter = ((AnnotationWriteField) writeField).getValueConverter()) != null) {
            val = valueConverter.convertValue(sheet, realRow, realCell, val);
        }
        return val;
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
    public void onWorkbookCreate(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void onSheetCreate(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }
    }

    @Override
    public void onCellValueSet(int c, int r, int s, Cell cell, Row row, Sheet sheet, Workbook workbook) {
        if (cell.getCellType() == CellType.STRING) {
            int length = cell.getStringCellValue().getBytes().length;
            WriteField writeField = writeFieldList.get(s).get(c);
            if (length > 0 && writeField.isAutoSize() && writeField.getWidth() < length) {
                writeField.setWidth(length);
            }
        }
    }

    @Override
    public void onSheetValueSet(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        for (int columnNum = 0; columnNum < writeFieldList.get(s).size(); columnNum++) {
            WriteField writeField = writeFieldList.get(s).get(columnNum);
            if (writeField.isAutoSize()) {
                sheet.autoSizeColumn(columnNum);
            }
            if (writeField.getWidth() > 0) {
                sheet.setColumnWidth(columnNum, writeField.getWidth() * 256 + 200);
            }
        }
    }

    public static class ListData {
        private String sheetName;
        private List<?> dataList;

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public List<?> getDataList() {
            return dataList;
        }

        public void setDataList(List<?> dataList) {
            this.dataList = dataList;
        }
    }
}
