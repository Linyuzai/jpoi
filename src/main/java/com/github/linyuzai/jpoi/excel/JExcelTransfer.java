package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.adapter.SimpleDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.adapter.TitleIndexDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.adapter.WriteAdapter;
import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.listener.PoiListener;
import com.github.linyuzai.jpoi.excel.setter.PoiValueSetter;
import com.github.linyuzai.jpoi.excel.setter.ValueSetter;
import com.github.linyuzai.jpoi.order.Ordered;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class JExcelTransfer {

    private Workbook workbook;
    private List<PoiListener> poiListeners;
    private List<ValueConverter> valueConverters;
    private ValueSetter valueSetter;
    private WriteAdapter writeAdapter;

    private boolean transferred = false;

    public JExcelTransfer(Workbook workbook) {
        this.workbook = workbook;
        this.poiListeners = new ArrayList<>();
        this.valueConverters = new ArrayList<>();
        setValueSetter(PoiValueSetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        addValueConverter(PictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(ObjectValueConverter.getInstance());
    }

    public JExcelTransfer(Workbook workbook, List<PoiListener> poiListeners, List<ValueConverter> valueConverters, ValueSetter valueSetter, WriteAdapter writeAdapter) {
        this.workbook = workbook;
        this.poiListeners = poiListeners;
        this.valueConverters = valueConverters;
        this.valueSetter = valueSetter;
        if (writeAdapter instanceof PoiListener && poiListeners != null) {
            poiListeners.remove(writeAdapter);
        }
    }

    public List<PoiListener> getPoiListeners() {
        return poiListeners;
    }

    public JExcelTransfer setPoiListeners(List<PoiListener> poiListeners) {
        this.poiListeners = poiListeners;
        this.poiListeners.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public void addPoiListener(PoiListener poiListener) {
        this.poiListeners.add(poiListener);
        this.poiListeners.sort(Comparator.comparingInt(Ordered::getOrder));
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public void setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
        this.valueConverters.sort(Comparator.comparingInt(Ordered::getOrder));
    }

    public JExcelTransfer addValueConverter(ValueConverter valueConverter) {
        this.valueConverters.add(valueConverter);
        this.valueConverters.sort(Comparator.comparingInt(Ordered::getOrder));
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

    public JExcelTransfer writeAdapter(WriteAdapter writeAdapter) {
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
        return writeAdapter(new SimpleDataWriteAdapter(data));
    }

    public JExcelTransfer data(List<?>... dataList) {
        return writeAdapter(new TitleIndexDataWriteAdapter(dataList));
    }

    public JExcelWriter write() {
        if (workbook == null) {
            throw new RuntimeException("No source to write");
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
            throw new RuntimeException("PoiListeners is null");
        }
        write(workbook, writeAdapter, poiListeners, valueConverters, valueSetter);
        return new JExcelWriter(workbook);
    }

    private void write(Workbook workbook, WriteAdapter writeAdapter, List<PoiListener> poiListeners, List<ValueConverter> valueConverters, ValueSetter valueSetter) {
        for (PoiListener poiListener : poiListeners) {
            poiListener.onWorkbookCreate(workbook);
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
                poiListener.onSheetCreate(s, sheet, drawing, workbook);
            }
            int rowCount = writeAdapter.getRowCount(s);
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.createRow(r);
                for (PoiListener poiListener : poiListeners) {
                    poiListener.onRowCreate(r, s, row, sheet, workbook);
                }
                int cellCount = writeAdapter.getCellCount(s, r);
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.createCell(c);
                    for (PoiListener poiListener : poiListeners) {
                        poiListener.onCellCreate(c, r, s, cell, row, sheet, workbook);
                    }
                    ValueConverter valueConverter = null;
                    Object o = writeAdapter.getData(s, r, c);
                    for (ValueConverter vc : valueConverters) {
                        if (vc.supportValue(s, r, c, o)) {
                            valueConverter = vc;
                            break;
                        }
                    }
                    if (valueConverter == null) {
                        throw new RuntimeException("No value converter matched");
                    }
                    Object value = valueConverter.adaptValue(s, r, c, o);
                    valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, value);
                }
            }
        }
        transferred = true;
    }

    public JExcelTransfer append() {
        return append(true);
    }

    public JExcelTransfer append(boolean exd) {
        if (!transferred) {
            write();
        }
        if (exd) {
            return new JExcelTransfer(workbook, poiListeners, valueConverters, valueSetter, writeAdapter);
        } else {
            return new JExcelTransfer(workbook);
        }
    }
}
