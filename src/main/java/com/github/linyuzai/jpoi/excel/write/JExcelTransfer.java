package com.github.linyuzai.jpoi.excel.write;

import com.github.linyuzai.jpoi.excel.JExcelBase;
import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.listener.PoiListener;
import com.github.linyuzai.jpoi.excel.write.adapter.SimpleDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.write.adapter.TitleIndexDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.write.adapter.WriteAdapter;
import com.github.linyuzai.jpoi.excel.write.auto.AutoWorkbook;
import com.github.linyuzai.jpoi.excel.write.setter.CombinationValueSetter;
import com.github.linyuzai.jpoi.excel.write.setter.ValueSetter;
import com.github.linyuzai.jpoi.support.SupportOrder;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class JExcelTransfer extends JExcelBase<JExcelTransfer> {

    private Workbook workbook;
    private List<PoiListener> poiListeners;
    private List<ValueConverter> valueConverters;
    private ValueSetter valueSetter;
    private WriteAdapter writeAdapter;

    public JExcelTransfer(Workbook workbook) {
        this.workbook = workbook;
        this.poiListeners = new ArrayList<>();
        this.valueConverters = new ArrayList<>();
        setValueSetter(CombinationValueSetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        //addValueConverter(WritePictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(BaseTypeValueConverter.getInstance());
        addValueConverter(WriteCombinationValueConverter.getInstance());
        addValueConverter(WriteSupportValueConverter.getInstance());
        addValueConverter(WriteObjectValueConverter.getInstance());
    }

    public List<PoiListener> getPoiListeners() {
        return poiListeners;
    }

    public JExcelTransfer setPoiListeners(List<PoiListener> poiListeners) {
        this.poiListeners = poiListeners;
        this.poiListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public JExcelTransfer addPoiListener(PoiListener poiListener) {
        this.poiListeners.add(poiListener);
        this.poiListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public JExcelTransfer setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
        this.valueConverters.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public JExcelTransfer addValueConverter(ValueConverter valueConverter) {
        this.valueConverters.add(valueConverter);
        this.valueConverters.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public ValueSetter getValueSetter() {
        return valueSetter;
    }

    public JExcelTransfer setValueSetter(ValueSetter valueSetter) {
        if (this.valueSetter instanceof PoiListener) {
            poiListeners.remove(this.valueSetter);
        }
        this.valueSetter = valueSetter;
        if (valueSetter instanceof PoiListener) {
            addPoiListener((PoiListener) valueSetter);
        }
        return this;
    }

    public WriteAdapter getWriteAdapter() {
        return writeAdapter;
    }

    public JExcelTransfer setWriteAdapter(WriteAdapter writeAdapter) {
        if (this.writeAdapter instanceof PoiListener) {
            poiListeners.remove(this.writeAdapter);
        }
        this.writeAdapter = writeAdapter;
        if (writeAdapter instanceof PoiListener) {
            addPoiListener((PoiListener) writeAdapter);
        }
        return this;
    }

    public JExcelTransfer data(List<?> data) {
        return setWriteAdapter(new SimpleDataWriteAdapter(data));
    }

    public JExcelTransfer data(List<?>... dataList) {
        return setWriteAdapter(new TitleIndexDataWriteAdapter(dataList));
    }

    public JExcelWriter write() {
        if (workbook == null) {
            throw new RuntimeException("No source to transfer");
        }
        if (writeAdapter == null) {
            throw new RuntimeException("WriteAdapter is null");
        }
        if (valueConverters == null) {
            throw new RuntimeException("ValueConverter is null");
        }
        if (valueSetter == null) {
            throw new RuntimeException("ValueSetter is null");
        }
        if (poiListeners == null) {
            throw new RuntimeException("PoiWriteListeners is null");
        }
        Workbook real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }
        transfer(real, writeAdapter, poiListeners, valueConverters, valueSetter);
        return new JExcelWriter(real);
    }

    private static void transfer(Workbook workbook, WriteAdapter writeAdapter, List<PoiListener> poiListeners,
                                 List<ValueConverter> valueConverters, ValueSetter valueSetter) {
        /*real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }*/
        CreationHelper creationHelper = workbook.getCreationHelper();
        for (PoiListener poiListener : poiListeners) {
            poiListener.onWorkbookStart(workbook, creationHelper);
        }
        int sheetCount = writeAdapter.getSheetCount();
        for (int s = 0; s < sheetCount; s++) {
            String sheetName = writeAdapter.getSheetName(s);
            Sheet sheet;
            if (sheetName == null) {
                sheet = workbook.createSheet();
            } else {
                sheet = workbook.createSheet(sheetName);
            }
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            for (PoiListener poiListener : poiListeners) {
                poiListener.onSheetStart(s, sheet, drawing, workbook);
            }
            int rowCount = writeAdapter.getRowCount(s);
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.createRow(r);
                for (PoiListener poiListener : poiListeners) {
                    poiListener.onRowStart(r, s, row, sheet, workbook);
                }
                int cellCount = writeAdapter.getCellCount(s, r);
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.createCell(c);
                    for (PoiListener poiListener : poiListeners) {
                        poiListener.onCellStart(c, r, s, cell, row, sheet, workbook);
                    }
                    Object o = writeAdapter.getData(s, r, c);
                    Object value = convertValue(valueConverters, s, r, c, o);
                    valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
                    for (PoiListener poiListener : poiListeners) {
                        poiListener.onCellEnd(c, r, s, cell, row, sheet, workbook);
                    }
                }
                for (PoiListener poiListener : poiListeners) {
                    poiListener.onRowEnd(r, s, row, sheet, workbook);
                }
            }
            for (PoiListener poiListener : poiListeners) {
                poiListener.onSheetEnd(s, sheet, drawing, workbook);
            }
        }
        for (PoiListener poiListener : poiListeners) {
            poiListener.onWorkbookEnd(workbook, creationHelper);
        }
    }
}
