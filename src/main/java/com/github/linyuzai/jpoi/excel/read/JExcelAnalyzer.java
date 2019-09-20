package com.github.linyuzai.jpoi.excel.read;

import com.github.linyuzai.jpoi.excel.JExcelBase;
import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.read.adapter.DirectListReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.MapReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.ObjectReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.ReadAdapter;
import com.github.linyuzai.jpoi.excel.read.getter.SupportValueGetter;
import com.github.linyuzai.jpoi.excel.read.getter.ValueGetter;
import com.github.linyuzai.jpoi.excel.read.listener.PoiReadListener;
import com.github.linyuzai.jpoi.order.Ordered;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class JExcelAnalyzer extends JExcelBase<JExcelAnalyzer> {

    private Workbook workbook;
    private List<PoiReadListener> poiReadListeners;
    private List<ValueConverter> valueConverters;
    private ValueGetter valueGetter;
    private ReadAdapter readAdapter;

    public JExcelAnalyzer(Workbook workbook) {
        this.workbook = workbook;
        this.poiReadListeners = new ArrayList<>();
        this.valueConverters = new ArrayList<>();
        setValueGetter(SupportValueGetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        //addValueConverter(PictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(ReadSupportValueConverter.getInstance());
        addValueConverter(ReadObjectValueConverter.getInstance());
    }

    public List<PoiReadListener> getPoiReadListeners() {
        return poiReadListeners;
    }

    public JExcelAnalyzer setPoiReadListeners(List<PoiReadListener> poiReadListeners) {
        this.poiReadListeners = poiReadListeners;
        this.poiReadListeners.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public JExcelAnalyzer addPoiReadListener(PoiReadListener poiReadListener) {
        this.poiReadListeners.add(poiReadListener);
        this.poiReadListeners.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public JExcelAnalyzer setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
        this.valueConverters.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public JExcelAnalyzer addValueConverter(ValueConverter valueConverter) {
        this.valueConverters.add(valueConverter);
        this.valueConverters.sort(Comparator.comparingInt(Ordered::getOrder));
        return this;
    }

    public ValueGetter getValueSetter() {
        return valueGetter;
    }

    public JExcelAnalyzer setValueGetter(ValueGetter valueGetter) {
        if (this.valueGetter instanceof PoiReadListener) {
            poiReadListeners.remove(this.valueGetter);
        }
        this.valueGetter = valueGetter;
        if (valueGetter instanceof PoiReadListener) {
            addPoiReadListener((PoiReadListener) valueGetter);
        }
        return this;
    }

    public ReadAdapter getReadAdapter() {
        return readAdapter;
    }

    public JExcelAnalyzer setReadAdapter(ReadAdapter readAdapter) {
        if (this.readAdapter instanceof PoiReadListener) {
            poiReadListeners.remove(this.readAdapter);
        }
        this.readAdapter = readAdapter;
        if (readAdapter instanceof PoiReadListener) {
            addPoiReadListener((PoiReadListener) readAdapter);
        }
        return this;
    }

    public JExcelAnalyzer target(Class<?>... classes) {
        setReadAdapter(new ObjectReadAdapter(classes));
        return this;
    }

    public JExcelAnalyzer toMap() {
        setReadAdapter(new MapReadAdapter());
        return this;
    }

    public JExcelAnalyzer direct() {
        setReadAdapter(new DirectListReadAdapter());
        return this;
    }

    public JExcelReader read() {
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
        if (poiReadListeners == null) {
            throw new RuntimeException("PoiReadListeners is null");
        }
        return new JExcelReader(analyze(workbook, readAdapter, valueConverters, valueGetter));
    }

    private Object analyze(Workbook workbook, ReadAdapter readAdapter, List<ValueConverter> valueConverters, ValueGetter valueGetter) {
        int sCount = workbook.getNumberOfSheets();
        for (int s = 0; s < sCount; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            int rCount = sheet.getLastRowNum() + 1;
            for (int r = 0; r < rCount; r++) {
                Row row = sheet.getRow(r);
                int cCount = row.getLastCellNum();
                for (int c = 0; c < cCount; c++) {
                    Cell cell = row.getCell(c);
                    Object o = valueGetter.getValue(s, r, c, cell, row, sheet, sheet.getDrawingPatriarch(), workbook);
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
                    Object cellValue = valueConverter.convertValue(s, r, c, o);
                    readAdapter.readCell(cellValue, s, r, c, sCount, rCount, cCount);
                }
            }
        }
        return readAdapter.getValue();
    }
}
