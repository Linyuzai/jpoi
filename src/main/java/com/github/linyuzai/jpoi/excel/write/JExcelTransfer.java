package com.github.linyuzai.jpoi.excel.write;

import com.github.linyuzai.jpoi.excel.cache.CacheManager;
import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.converter.*;
import com.github.linyuzai.jpoi.excel.handler.ExcelExceptionHandler;
import com.github.linyuzai.jpoi.excel.handler.ExceptionValue;
import com.github.linyuzai.jpoi.excel.listener.ExcelListener;
import com.github.linyuzai.jpoi.excel.post.PostProcessor;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;
import com.github.linyuzai.jpoi.excel.write.adapter.SimpleDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.write.adapter.TitleIndexDataWriteAdapter;
import com.github.linyuzai.jpoi.excel.write.adapter.WriteAdapter;
import com.github.linyuzai.jpoi.excel.write.auto.AutoWorkbook;
import com.github.linyuzai.jpoi.excel.write.setter.CombinationValueSetter;
import com.github.linyuzai.jpoi.excel.write.setter.ValueSetter;
import com.github.linyuzai.jpoi.exception.JPoiException;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class JExcelTransfer extends JExcelProcessor<JExcelTransfer> {

    private ValueSetter valueSetter;
    private WriteAdapter writeAdapter;

    public JExcelTransfer(Workbook workbook) {
        super(workbook);
        setValueSetter(CombinationValueSetter.getInstance());
        addValueConverter(NullValueConverter.getInstance());
        addValueConverter(PostValueConverter.getInstance());
        //addValueConverter(WritePictureValueConverter.getInstance());
        addValueConverter(PoiValueConverter.getInstance());
        addValueConverter(BaseTypeValueConverter.getInstance());
        addValueConverter(WriteCombinationValueConverter.getInstance());
        addValueConverter(WriteSupportValueConverter.getInstance());
        addValueConverter(WriteObjectValueConverter.getInstance());
    }

    public ValueSetter getValueSetter() {
        return valueSetter;
    }

    public JExcelTransfer setValueSetter(ValueSetter valueSetter) {
        if (this.valueSetter instanceof ExcelListener) {
            excelListeners.remove(this.valueSetter);
        }
        this.valueSetter = valueSetter;
        if (valueSetter instanceof ExcelListener) {
            addListener((ExcelListener) valueSetter);
        }
        return this;
    }

    public WriteAdapter getWriteAdapter() {
        return writeAdapter;
    }

    public JExcelTransfer setWriteAdapter(WriteAdapter writeAdapter) {
        if (this.writeAdapter instanceof ExcelListener) {
            excelListeners.remove(this.writeAdapter);
        }
        this.writeAdapter = writeAdapter;
        if (writeAdapter instanceof ExcelListener) {
            addListener((ExcelListener) writeAdapter);
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
        check();
        if (writeAdapter == null) {
            throw new JPoiException("WriteAdapter is null");
        }
        if (valueSetter == null) {
            throw new JPoiException("ValueSetter is null");
        }
        Workbook real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }
        return new JExcelWriter(transfer(real, writeAdapter, excelListeners, valueConverters,
                valueSetter, postProcessor, cacheManager, excelExceptionHandler));
    }

    private Values transfer(Workbook workbook, WriteAdapter writeAdapter, List<ExcelListener> excelListeners,
                            List<ValueConverter> valueConverters, ValueSetter valueSetter,
                            PostProcessor postProcessor, CacheManager cacheManager, ExcelExceptionHandler exceptionHandler) {
        /*real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }*/
        boolean forceBreak = false;
        List<PostValue> postValues = new ArrayList<>();
        List<ExceptionValue> exceptionValues = new ArrayList<>();
        for (int w = 0; w < 1; w++) {
            CreationHelper creationHelper = workbook.getCreationHelper();
            for (ExcelListener excelListener : excelListeners) {
                excelListener.onWorkbookStart(workbook, creationHelper);
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
                for (ExcelListener excelListener : excelListeners) {
                    excelListener.onSheetStart(s, sheet, drawing, workbook);
                }
                int rowCount = writeAdapter.getRowCount(s);
                for (int r = 0; r < rowCount; r++) {
                    Row row = sheet.createRow(r);
                    for (ExcelListener excelListener : excelListeners) {
                        excelListener.onRowStart(r, s, row, sheet, workbook);
                    }
                    int cellCount = writeAdapter.getCellCount(s, r);
                    for (int c = 0; c < cellCount; c++) {
                        Cell cell = row.createCell(c);
                        for (ExcelListener excelListener : excelListeners) {
                            excelListener.onCellStart(c, r, s, cell, row, sheet, workbook);
                        }
                        try {
                            Object source = writeAdapter.getData(s, r, c);
                            Object cache = cacheManager.getCache(this, source, s, r, c);
                            Object value;
                            if (cache == null) {
                                value = convertValue(valueConverters, s, r, c, source);
                            } else {
                                value = cache;
                            }
                            if (value instanceof PostValue) {
                                PostValue postValue = (PostValue) value;
                                fillPostValue(postValue, w, s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
                                Object postCache = cacheManager.getCache(this, postValue, s, r, c);
                                if (postCache == null) {
                                    postValues.add(postValue);
                                } else {
                                    valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, postCache);
                                }
                            } else {
                                if (cache == null) {
                                    cacheManager.setCache(this, source, value, s, r, c);
                                }
                                valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
                            }
                        } catch (Throwable e) {
                            exceptionValues.add(new ExceptionValue(s, r, c, e));
                            forceBreak = exceptionHandler.handle(this, s, r, c, cell, row, sheet, workbook, e);
                        }
                        if (forceBreak) {
                            break;
                        }
                        for (ExcelListener excelListener : excelListeners) {
                            excelListener.onCellEnd(c, r, s, cell, row, sheet, workbook);
                        }
                    }
                    if (forceBreak) {
                        break;
                    }
                    for (ExcelListener excelListener : excelListeners) {
                        excelListener.onRowEnd(r, s, row, sheet, workbook);
                    }
                }
                if (forceBreak) {
                    break;
                }
                for (ExcelListener excelListener : excelListeners) {
                    excelListener.onSheetEnd(s, sheet, drawing, workbook);
                }
            }
            if (forceBreak) {
                break;
            }
            for (ExcelListener excelListener : excelListeners) {
                excelListener.onWorkbookEnd(workbook, creationHelper);
            }
        }
        if (!forceBreak) {
            exceptionValues.addAll(postProcessor.processPost(postValues, this));
            for (PostValue pv : postValues) {
                int s = pv.getSheetIndex();
                int r = pv.getRowIndex();
                int c = pv.getCellIndex();
                Cell cell = pv.getCell();
                Row row = pv.getRow();
                Sheet sheet = pv.getSheet();
                Object value = pv.getValue();
                try {
                    cacheManager.setCache(this, pv, value, s, r, c);
                    valueSetter.setValue(s, r, c, cell, row, sheet, pv.getDrawing(), workbook, pv.getCreationHelper(), value);
                } catch (Throwable e) {
                    exceptionValues.add(new ExceptionValue(s, r, c, e));
                    forceBreak = exceptionHandler.handle(this, s, r, c, cell, row, sheet, workbook, e);
                }
                if (forceBreak) {
                    break;
                }
            }
        }
        return new Values(workbook, exceptionValues);
    }

    public static class Values {
        private Workbook workbook;
        private List<ExceptionValue> exceptionValues;

        public Values(Workbook workbook, List<ExceptionValue> exceptionValues) {
            this.workbook = workbook;
            this.exceptionValues = exceptionValues;
        }

        public Workbook getWorkbook() {
            return workbook;
        }

        public void setWorkbook(Workbook workbook) {
            this.workbook = workbook;
        }

        public List<ExceptionValue> getExceptionValues() {
            return exceptionValues;
        }

        public void setExceptionValues(List<ExceptionValue> exceptionValues) {
            this.exceptionValues = exceptionValues;
        }
    }
}
