package com.github.linyuzai.jpoi.excel.read;

import com.github.linyuzai.jpoi.excel.JExcelBase;
import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.listener.PoiListener;
import com.github.linyuzai.jpoi.excel.read.adapter.DirectListReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.MapReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.ObjectReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.ReadAdapter;
import com.github.linyuzai.jpoi.excel.read.getter.CombinationValueGetter;
import com.github.linyuzai.jpoi.excel.read.getter.ValueGetter;
import com.github.linyuzai.jpoi.support.SupportOrder;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class JExcelAnalyzer extends JExcelBase<JExcelAnalyzer> {

    private Workbook workbook;
    private List<PoiListener> poiListeners;
    private List<ValueConverter> valueConverters;
    private ValueGetter valueGetter;
    private ReadAdapter readAdapter;

    public JExcelAnalyzer(Workbook workbook) {
        this.workbook = workbook;
        this.poiListeners = new ArrayList<>();
        this.valueConverters = new ArrayList<>();
        setValueGetter(CombinationValueGetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        //addValueConverter(WritePictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(ReadSupportValueConverter.getInstance());
        //addValueConverter(ReadDataValueConverter.getInstance());
        addValueConverter(ReadObjectValueConverter.getInstance());
    }

    public List<PoiListener> getPoiListeners() {
        return poiListeners;
    }

    public JExcelAnalyzer setPoiListeners(List<PoiListener> poiListeners) {
        this.poiListeners = poiListeners;
        this.poiListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public JExcelAnalyzer addPoiReadListener(PoiListener poiListener) {
        this.poiListeners.add(poiListener);
        this.poiListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public JExcelAnalyzer setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
        this.valueConverters.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public JExcelAnalyzer addValueConverter(ValueConverter valueConverter) {
        this.valueConverters.add(valueConverter);
        this.valueConverters.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return this;
    }

    public ValueGetter getValueSetter() {
        return valueGetter;
    }

    public JExcelAnalyzer setValueGetter(ValueGetter valueGetter) {
        if (this.valueGetter instanceof PoiListener) {
            poiListeners.remove(this.valueGetter);
        }
        this.valueGetter = valueGetter;
        if (valueGetter instanceof PoiListener) {
            addPoiReadListener((PoiListener) valueGetter);
        }
        return this;
    }

    public ReadAdapter getReadAdapter() {
        return readAdapter;
    }

    public JExcelAnalyzer setReadAdapter(ReadAdapter readAdapter) {
        if (this.readAdapter instanceof PoiListener) {
            poiListeners.remove(this.readAdapter);
        }
        this.readAdapter = readAdapter;
        if (readAdapter instanceof PoiListener) {
            addPoiReadListener((PoiListener) readAdapter);
        }
        return this;
    }

    public JExcelAnalyzer target(Class<?>... classes) {
        return setReadAdapter(new ObjectReadAdapter(classes));
    }

    public JExcelAnalyzer toMap() {
        return setReadAdapter(new MapReadAdapter());
    }

    public JExcelAnalyzer direct() {
        return setReadAdapter(new DirectListReadAdapter());
    }

    public JExcelReader read() throws IOException {
        return read(true);
    }

    public JExcelReader read(boolean close) throws IOException {
        if (workbook == null) {
            throw new RuntimeException("No source to transfer");
        }
        if (readAdapter == null) {
            throw new RuntimeException("ReadAdapter is null");
        }
        if (valueConverters == null) {
            throw new RuntimeException("ValueConverter is null");
        }
        if (valueGetter == null) {
            throw new RuntimeException("ValueGetter is null");
        }
        if (poiListeners == null) {
            throw new RuntimeException("PoiReadListeners is null");
        }
        return new JExcelReader(analyze(workbook, readAdapter, poiListeners, valueConverters, valueGetter, close));
    }

    private static Object analyze(Workbook workbook, ReadAdapter readAdapter, List<PoiListener> poiListeners,
                                  List<ValueConverter> valueConverters, ValueGetter valueGetter, boolean close) throws IOException {
        CreationHelper creationHelper = workbook.getCreationHelper();
        for (PoiListener poiListener : poiListeners) {
            poiListener.onWorkbookStart(workbook, creationHelper);
        }
        int sCount = workbook.getNumberOfSheets();
        for (int s = 0; s < sCount; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            Drawing<?> drawing = sheet.getDrawingPatriarch();
            for (PoiListener poiListener : poiListeners) {
                poiListener.onSheetStart(s, sheet, drawing, workbook);
            }
            int rCount = sheet.getLastRowNum() + 1;
            for (int r = 0; r < rCount; r++) {
                Row row = sheet.getRow(r);
                for (PoiListener poiListener : poiListeners) {
                    poiListener.onRowStart(r, s, row, sheet, workbook);
                }
                int cCount = row.getLastCellNum();
                for (int c = 0; c < cCount; c++) {
                    Cell cell = row.getCell(c);
                    for (PoiListener poiListener : poiListeners) {
                        poiListener.onCellStart(c, r, s, cell, row, sheet, workbook);
                    }
                    Object o = valueGetter.getValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
                    /*ValueConverter valueConverter = null;
                    for (ValueConverter vc : valueConverters) {
                        if (vc.supportValue(s, r, c, o)) {
                            valueConverter = vc;
                            break;
                        }
                    }
                    if (valueConverter == null) {
                        throw new RuntimeException("No value converter matched");
                    }
                    Object cellValue = valueConverter.convertValue(s, r, c, o);*/
                    Object cellValue = convertValue(valueConverters, s, r, c, o);
                    readAdapter.readCell(cellValue, s, r, c, sCount, rCount, cCount);
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
        if (close) {
            workbook.close();
        }
        return readAdapter.getValue();
    }
}
