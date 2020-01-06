package com.github.linyuzai.jpoi.excel.read;

import com.github.linyuzai.jpoi.cache.CacheManager;
import com.github.linyuzai.jpoi.excel.JExcelBase;
import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.handler.ExcelExceptionHandler;
import com.github.linyuzai.jpoi.excel.listener.ExcelListener;
import com.github.linyuzai.jpoi.excel.processor.PostProcessor;
import com.github.linyuzai.jpoi.excel.read.adapter.DirectListReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.MapReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.ObjectReadAdapter;
import com.github.linyuzai.jpoi.excel.read.adapter.ReadAdapter;
import com.github.linyuzai.jpoi.excel.read.getter.CombinationValueGetter;
import com.github.linyuzai.jpoi.excel.read.getter.ValueGetter;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;
import com.github.linyuzai.jpoi.exception.JPoiException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class JExcelAnalyzer extends JExcelBase<JExcelAnalyzer> {

    private ValueGetter valueGetter;
    private ReadAdapter readAdapter;

    public JExcelAnalyzer(Workbook workbook) {
        super(workbook);
        setValueGetter(CombinationValueGetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        addValueConverter(ReadFormulaValueConverter.getInstance());
        addValueConverter(PostValueConverter.getInstance());
        //addValueConverter(WritePictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(ReadSupportValueConverter.getInstance());
        //addValueConverter(ReadDataValueConverter.getInstance());
        addValueConverter(ReadObjectValueConverter.getInstance());
    }

    public ValueGetter getValueSetter() {
        return valueGetter;
    }

    public JExcelAnalyzer setValueGetter(ValueGetter valueGetter) {
        if (this.valueGetter instanceof ExcelListener) {
            excelListeners.remove(this.valueGetter);
        }
        this.valueGetter = valueGetter;
        if (valueGetter instanceof ExcelListener) {
            addListener((ExcelListener) valueGetter);
        }
        return this;
    }

    public ReadAdapter getReadAdapter() {
        return readAdapter;
    }

    public JExcelAnalyzer setReadAdapter(ReadAdapter readAdapter) {
        if (this.readAdapter instanceof ExcelListener) {
            excelListeners.remove(this.readAdapter);
        }
        this.readAdapter = readAdapter;
        if (readAdapter instanceof ExcelListener) {
            addListener((ExcelListener) readAdapter);
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
        check();
        if (readAdapter == null) {
            throw new JPoiException("ReadAdapter is null");
        }
        if (valueGetter == null) {
            throw new JPoiException("ValueGetter is null");
        }
        return new JExcelReader(analyze(workbook, readAdapter, excelListeners, valueConverters,
                valueGetter, postProcessor, cacheManager, excelExceptionHandler, close));
    }

    private Values analyze(Workbook workbook, ReadAdapter readAdapter, List<ExcelListener> listeners,
                           List<ValueConverter> valueConverters, ValueGetter valueGetter,
                           PostProcessor postProcessor, CacheManager cacheManager,
                           ExcelExceptionHandler exceptionHandler, boolean close) throws IOException {
        boolean forceBreak = false;
        List<PostValue> postValues = new ArrayList<>();
        List<Throwable> throwableRecords = new ArrayList<>();
        for (int w = 0; w < 1; w++) {
            CreationHelper creationHelper = workbook.getCreationHelper();
            for (ExcelListener excelListener : listeners) {
                excelListener.onWorkbookStart(workbook, creationHelper);
            }
            int sCount = workbook.getNumberOfSheets();
            for (int s = 0; s < sCount; s++) {
                Sheet sheet = workbook.getSheetAt(s);
                if (sheet == null) {
                    continue;
                }
                Drawing<?> drawing = sheet.getDrawingPatriarch();
                for (ExcelListener excelListener : listeners) {
                    excelListener.onSheetStart(s, sheet, drawing, workbook);
                }
                int rCount = sheet.getLastRowNum() + 1;
                for (int r = 0; r < rCount; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }
                    for (ExcelListener excelListener : listeners) {
                        excelListener.onRowStart(r, s, row, sheet, workbook);
                    }
                    int cCount = row.getLastCellNum();
                    for (int c = 0; c < cCount; c++) {
                        Cell cell = row.getCell(c);
                        if (cell == null) {
                            continue;
                        }
                        for (ExcelListener excelListener : listeners) {
                            excelListener.onCellStart(c, r, s, cell, row, sheet, workbook);
                        }
                        try {
                            Object source = valueGetter.getValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
                            Object cache = cacheManager.getCache(this, source, s, r, c);
                            Object value;
                            if (cache == null) {
                                value = convertValue(valueConverters, s, r, c, source);
                            } else {
                                value = cache;
                            }
                            Object returnValue = readAdapter.readCell(value, s, r, c, sCount, rCount, cCount);
                            if (returnValue instanceof PostValue) {
                                PostValue postValue = (PostValue) returnValue;
                                fillPostValue(postValue, w, s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
                                Object postCache = cacheManager.getCache(this, postValue, s, r, c);
                                if (postCache == null) {
                                    postValues.add(postValue);
                                } else {
                                    readAdapter.readCell(postCache, s, r, c, sCount, rCount, cCount);
                                }
                            }
                        } catch (Throwable e) {
                            throwableRecords.add(e);
                            forceBreak = exceptionHandler.handle(s, r, c, cell, row, sheet, workbook, e);
                        }
                        if (forceBreak) {
                            break;
                        }
                        for (ExcelListener excelListener : listeners) {
                            excelListener.onCellEnd(c, r, s, cell, row, sheet, workbook);
                        }
                    }
                    if (forceBreak) {
                        break;
                    }
                    for (ExcelListener excelListener : listeners) {
                        excelListener.onRowEnd(r, s, row, sheet, workbook);
                    }
                }
                if (forceBreak) {
                    break;
                }
                for (ExcelListener excelListener : listeners) {
                    excelListener.onSheetEnd(s, sheet, drawing, workbook);
                }
            }
            if (forceBreak) {
                break;
            }
            for (ExcelListener excelListener : listeners) {
                excelListener.onWorkbookEnd(workbook, creationHelper);
            }
        }
        if (!forceBreak) {
            throwableRecords.addAll(postProcessor.processPost(postValues, exceptionHandler));
            for (PostValue pv : postValues) {
                int s = pv.getSheetIndex();
                int r = pv.getRowIndex();
                int c = pv.getCellIndex();
                Object value = pv.getValue();
                try {
                    cacheManager.setCache(this, pv, value, s, r, c);
                    readAdapter.readCell(value, s, r, c, -1, -1, -1);
                } catch (Throwable e) {
                    throwableRecords.add(e);
                    forceBreak = exceptionHandler.handle(s, r, c, pv.getCell(), pv.getRow(), pv.getSheet(), workbook, e);
                }
                if (forceBreak) {
                    break;
                }
            }
        }
        if (close) {
            workbook.close();
        }
        return new Values(readAdapter.getValue(), throwableRecords);
    }

    public static class Values {
        private Object value;
        private List<Throwable> throwableRecords;

        public Values(Object value, List<Throwable> throwableRecords) {
            this.value = value;
            this.throwableRecords = throwableRecords;
        }

        public Object getValue() {
            return value;
        }

        public void setValue(Object value) {
            this.value = value;
        }

        public List<Throwable> getThrowableRecords() {
            return throwableRecords;
        }

        public void setThrowableRecords(List<Throwable> throwableRecords) {
            this.throwableRecords = throwableRecords;
        }
    }
}
