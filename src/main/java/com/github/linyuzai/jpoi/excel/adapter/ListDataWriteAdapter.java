package com.github.linyuzai.jpoi.excel.adapter;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.listener.PoiListener;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

public class ListDataWriteAdapter extends AnnotationWriteAdapter implements PoiListener {

    private List<ListData> listDataList = new ArrayList<>();
    private List<List<FieldData>> fieldDataList = new ArrayList<>();

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

    public List<List<FieldData>> getFieldDataList() {
        return fieldDataList;
    }

    public void cellFields(int sheet, List<?> dataList, Class<?> cls) {
        FieldData fieldData = getFieldDataIncludeAnnotation(cls);
        if (fieldData != null) {
            /*StringBuilder sheetName = new StringBuilder(fieldData.getFieldDescription());
            for (ListData listData : listDataList) {
                if (sheetName.toString().equals(listData.getSheetName())) {
                    sheetName.append(sheet);
                    break;
                }
            }*/
            listDataList.get(sheet).setSheetName(fieldData.getFieldDescription());
        }
        if (cls != null && !cls.isInterface() && !cls.isArray() && !cls.isPrimitive() && !isWrapClass(cls) && cls != String.class) {
            fieldDataList.add(sheet, getFieldDataList(cls));
        } else {
            if (dataList.size() > 0) {
                //fieldDataList.add(sheet, getFieldDataList(dataList.get(0).getClass()));
                cellFields(sheet, dataList, dataList.get(0).getClass());
            } else {
                fieldDataList.add(sheet, Collections.emptyList());
            }
        }
        for (List<FieldData> fd : fieldDataList) {
            fd.sort(Comparator.comparingInt(FieldData::getOrder));
        }
    }

    public List<FieldData> getFieldDataList(Class<?> cls) {
        if (illegalFieldDataClass(cls)) {
            return Collections.emptyList();
        }
        Field[] fields = cls.getDeclaredFields();
        List<FieldData> fieldDataList = new ArrayList<>();
        for (Field field : fields) {
            FieldData fieldData = getFieldDataIncludeAnnotation(field);
            if (isAnnotationOnly()) {
                if (fieldData instanceof AnnotationFieldData) {
                    fieldDataList.add(fieldData);
                }
            } else {
                fieldDataList.add(fieldData);
            }
        }
        Method[] methods = cls.getDeclaredMethods();
        for (Method method : methods) {
            FieldData fieldData = getFieldDataIncludeAnnotation(method);
            if (fieldData instanceof AnnotationFieldData) {
                fieldDataList.add(fieldData);
            }
        }
        return fieldDataList;
    }

    public boolean illegalFieldDataClass(Class<?> cls) {
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
        return fieldDataList.get(sheet).size();
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
        FieldData fieldData = fieldDataList.get(sheet).get(cell);
        String fieldName = fieldData.getFieldName();
        Object val;
        if (fieldData instanceof AnnotationFieldData && ((AnnotationFieldData) fieldData).isMethod()) {
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
        if (fieldData instanceof AnnotationFieldData &&
                (valueConverter = ((AnnotationFieldData) fieldData).getValueConverter()) != null) {
            val = valueConverter.adaptValue(sheet, realRow, realCell, val);
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
    public void onSheetValueSet(int s, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        for (int columnNum = 0; columnNum < fieldDataList.size(); columnNum++) {
            if (fieldDataList.get(s).get(columnNum).isAutoSize()) {
                int columnWidth = sheet.getColumnWidth(columnNum) / 256;
                for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                    Row currentRow = sheet.getRow(rowNum);
                    /*if (sheet.getRow(rowNum) == null) {
                        currentRow = sheet.createRow(rowNum);
                    }*/
                    if (currentRow != null && currentRow.getCell(columnNum) != null) {
                        Cell currentCell = currentRow.getCell(columnNum);
                        if (currentCell.getCellType() == CellType.STRING) {
                            int length = currentCell.getStringCellValue().getBytes().length;
                            if (columnWidth < length) {
                                columnWidth = length;
                            }
                        }
                    }
                }
                sheet.setColumnWidth(columnNum, columnWidth * 256);
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
