package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.cache.CacheManager;
import com.github.linyuzai.jpoi.excel.cache.UncachedCacheManager;
import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.handler.ExcelExceptionHandler;
import com.github.linyuzai.jpoi.excel.handler.InterruptedExceptionHandler;
import com.github.linyuzai.jpoi.excel.listener.ExcelListener;
import com.github.linyuzai.jpoi.excel.processor.EmptyPostProcessor;
import com.github.linyuzai.jpoi.excel.processor.PostProcessor;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;
import com.github.linyuzai.jpoi.exception.JPoiException;
import com.github.linyuzai.jpoi.support.SupportOrder;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@SuppressWarnings("unchecked")
public abstract class JExcelBase<T extends JExcelBase<T>> {

    protected Workbook workbook;
    protected List<ExcelListener> excelListeners;
    protected List<ValueConverter> valueConverters;
    protected PostProcessor postProcessor;
    protected CacheManager cacheManager;
    protected ExcelExceptionHandler excelExceptionHandler;

    public JExcelBase(Workbook workbook) {
        this.workbook = workbook;
        this.excelListeners = new ArrayList<>();
        this.valueConverters = new ArrayList<>();
        this.postProcessor = EmptyPostProcessor.getInstance();
        this.cacheManager = UncachedCacheManager.getInstance();
        this.excelExceptionHandler = InterruptedExceptionHandler.getInstance();
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public List<ExcelListener> getListeners() {
        return excelListeners;
    }

    public T setListeners(List<ExcelListener> excelListeners) {
        this.excelListeners = excelListeners;
        this.excelListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return (T) this;
    }

    public T addListener(ExcelListener excelListener) {
        this.excelListeners.add(excelListener);
        this.excelListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return (T) this;
    }

    public T removeListener(ExcelListener excelListener) {
        this.excelListeners.remove(excelListener);
        this.excelListeners.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return (T) this;
    }

    public List<ValueConverter> getValueConverters() {
        return valueConverters;
    }

    public T setValueConverters(List<ValueConverter> valueConverters) {
        this.valueConverters = valueConverters;
        this.valueConverters.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return (T) this;
    }

    public T addValueConverter(ValueConverter valueConverter) {
        this.valueConverters.add(valueConverter);
        this.valueConverters.sort(Comparator.comparingInt(SupportOrder::getOrder));
        return (T) this;
    }

    public PostProcessor getPostProcessor() {
        return postProcessor;
    }

    public T setPostProcessor(PostProcessor postProcessor) {
        this.postProcessor = postProcessor;
        return (T) this;
    }

    public CacheManager getCacheManager() {
        return cacheManager;
    }

    public void setCacheManager(CacheManager cacheManager) {
        this.cacheManager = cacheManager;
    }

    public ExcelExceptionHandler getExceptionHandler() {
        return excelExceptionHandler;
    }

    public T setExceptionHandler(ExcelExceptionHandler excelExceptionHandler) {
        this.excelExceptionHandler = excelExceptionHandler;
        return (T) this;
    }

    public void check() {
        if (workbook == null) {
            throw new JPoiException("No source to transfer");
        }
        if (excelListeners == null) {
            throw new JPoiException("ExcelListeners is null");
        }
        if (valueConverters == null) {
            throw new JPoiException("ValueConverter is null");
        }
        if (postProcessor == null) {
            throw new JPoiException("PostProcessor is null");
        }
        if (cacheManager == null) {
            throw new JPoiException("CacheManager is null");
        }
        if (excelExceptionHandler == null) {
            throw new JPoiException("ExcelExceptionHandler is null");
        }
    }

    public static void fillPostValue(PostValue postValue, int w, int s, int r, int c, Cell cell, Row row, Sheet sheet,
                                     Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper) {
        postValue.setWorkbook(workbook);
        postValue.setSheet(sheet);
        postValue.setRow(row);
        postValue.setCell(cell);
        postValue.setWorkbookIndex(w);
        postValue.setSheetIndex(s);
        postValue.setRowIndex(r);
        postValue.setCellIndex(c);
        postValue.setDrawing(drawing);
        postValue.setCreationHelper(creationHelper);
    }

    public static Object convertValue(List<ValueConverter> valueConverters, int s, int r, int c, Object o) {
        ValueConverter valueConverter = null;
        for (ValueConverter vc : valueConverters) {
            if (vc.supportValue(s, r, c, o)) {
                valueConverter = vc;
                break;
            }
        }
        if (valueConverter == null) {
            throw new JPoiException("No value converter matched");
        }
        return valueConverter.convertValue(s, r, c, o);
    }
}
