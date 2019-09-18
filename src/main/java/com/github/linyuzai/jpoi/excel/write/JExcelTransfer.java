package com.github.linyuzai.jpoi.excel.write;

import com.github.linyuzai.jpoi.excel.write.adapter.SimpleDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.write.adapter.TitleIndexDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.write.adapter.WriteAdapter;
import com.github.linyuzai.jpoi.excel.write.workbook.AutoWorkbook;
import com.github.linyuzai.jpoi.excel.write.converter.*;
import com.github.linyuzai.jpoi.excel.write.listener.PoiWriteListener;
import com.github.linyuzai.jpoi.excel.write.setter.SupportValueSetter;
import com.github.linyuzai.jpoi.excel.write.setter.ValueSetter;
import com.github.linyuzai.jpoi.order.Ordered;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class JExcelTransfer {

    private Workbook workbook;
    private Workbook real;
    private List<PoiWriteListener> poiWriteListeners;
    private List<ValueConverter> valueConverters;
    private ValueSetter valueSetter;
    private WriteAdapter writeAdapter;

    private boolean transferred = false;

    public JExcelTransfer(Workbook workbook) {
        this.workbook = workbook;
        this.poiWriteListeners = new ArrayList<>();
        this.valueConverters = new ArrayList<>();
        setValueSetter(SupportValueSetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        //addValueConverter(PictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(SupportValueConverter.getInstance());
        addValueConverter(ObjectValueConverter.getInstance());
    }

    public JExcelTransfer(Workbook workbook, List<PoiWriteListener> poiWriteListeners, List<ValueConverter> valueConverters, ValueSetter valueSetter, WriteAdapter writeAdapter) {
        this.workbook = workbook;
        this.poiWriteListeners = poiWriteListeners;
        this.valueConverters = valueConverters;
        this.valueSetter = valueSetter;
        if (writeAdapter instanceof PoiWriteListener && poiWriteListeners != null) {
            poiWriteListeners.remove(writeAdapter);
        }
    }

    public List<PoiWriteListener> getPoiWriteListeners() {
        return poiWriteListeners;
    }

    public JExcelTransfer setPoiWriteListeners(List<PoiWriteListener> poiWriteListeners) {
        this.poiWriteListeners = poiWriteListeners;
        this.poiWriteListeners.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public JExcelTransfer addPoiListener(PoiWriteListener poiWriteListener) {
        this.poiWriteListeners.add(poiWriteListener);
        this.poiWriteListeners.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public JExcelTransfer setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
        this.valueConverters.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
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
        if (this.valueSetter instanceof PoiWriteListener) {
            poiWriteListeners.remove(this.valueSetter);
        }
        this.valueSetter = valueSetter;
        if (valueSetter instanceof PoiWriteListener) {
            addPoiListener((PoiWriteListener) valueSetter);
        }
        return this;
    }

    public WriteAdapter getWriteAdapter() {
        return writeAdapter;
    }

    public JExcelTransfer writeAdapter(WriteAdapter writeAdapter) {
        if (this.writeAdapter instanceof PoiWriteListener) {
            poiWriteListeners.remove(this.writeAdapter);
        }
        this.writeAdapter = writeAdapter;
        if (writeAdapter instanceof PoiWriteListener) {
            addPoiListener((PoiWriteListener) writeAdapter);
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
        if (poiWriteListeners == null) {
            throw new RuntimeException("PoiListeners is null");
        }
        transfer(workbook, writeAdapter, poiWriteListeners, valueConverters, valueSetter);
        return new JExcelWriter(real);
    }

    private void transfer(Workbook workbook, WriteAdapter writeAdapter, List<PoiWriteListener> poiWriteListeners, List<ValueConverter> valueConverters, ValueSetter valueSetter) {
        real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }
        for (PoiWriteListener poiWriteListener : poiWriteListeners) {
            poiWriteListener.onWorkbookCreate(real);
        }
        int sheetCount = writeAdapter.getSheetCount();
        for (int s = 0; s < sheetCount; s++) {
            String sheetName = writeAdapter.getSheetName(s);
            Sheet sheet;
            if (sheetName == null) {
                sheet = real.createSheet();
            } else {
                sheet = real.createSheet(sheetName);
            }
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetCreate(s, sheet, drawing, real);
            }
            int rowCount = writeAdapter.getRowCount(s);
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.createRow(r);
                for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                    poiWriteListener.onRowCreate(r, s, row, sheet, real);
                }
                int cellCount = writeAdapter.getCellCount(s, r);
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.createCell(c);
                    for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                        poiWriteListener.onCellCreate(c, r, s, cell, row, sheet, real);
                    }
                    Object o = writeAdapter.getData(s, r, c);
                    ValueConverter valueConverter = null;
                    for (ValueConverter vc : valueConverters) {
                        if (vc.supportValue(s, r, c, o)) {
                            valueConverter = vc;
                            break;
                        }
                    }
                    if (valueConverter == null) {
                        throw new RuntimeException("No value converter matched");
                    }
                    Object value = valueConverter.convertValue(s, r, c, o);
                    valueSetter.setValue(s, r, c, cell, row, sheet, drawing, real, value);
                    for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                        poiWriteListener.onCellValueSet(c, r, s, cell, row, sheet, real);
                    }
                }
                for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                    poiWriteListener.onRowValueSet(r, s, row, sheet, real);
                }
            }
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetValueSet(s, sheet, drawing, real);
            }
        }
        for (PoiWriteListener poiWriteListener : poiWriteListeners) {
            poiWriteListener.onWorkbookValueSet(real);
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
            return new JExcelTransfer(real, poiWriteListeners, valueConverters, valueSetter, writeAdapter);
        } else {
            return new JExcelTransfer(real);
        }
    }
}
