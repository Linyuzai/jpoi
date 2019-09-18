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
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

public class JExcelTransfer {

    private Workbook workbook;
    private Workbook real;
    private List<PoiWriteListener> poiWriteListeners;
    private List<ValueConverter> valueConverters;
    private ValueSetter valueSetter;
    private WriteAdapter writeAdapter;

    private Executor executor = Executors.newCachedThreadPool();
    private boolean async = false;
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

    @Deprecated
    public JExcelTransfer async() {
        async = true;
        return this;
    }

    public JExcelTransfer sync() {
        async = false;
        return this;
    }

    public boolean isAsync() {
        return async;
    }

    public JExcelTransfer executor(Executor executor) {
        this.executor = executor;
        return this;
    }

    public Executor getExecutor() {
        return executor;
    }

    public boolean isTransferred() {
        return transferred;
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
        real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }

        //long t1 = System.currentTimeMillis();

        if (async) {
            transferAsync(executor, real, writeAdapter, poiWriteListeners, valueConverters, valueSetter);
        } else {
            transfer(real, writeAdapter, poiWriteListeners, valueConverters, valueSetter);
        }
        //System.out.println(System.currentTimeMillis() - t1);
        transferred = true;
        return new JExcelWriter(real);
    }

    private static void transfer(Workbook workbook, WriteAdapter writeAdapter, List<PoiWriteListener> poiWriteListeners, List<ValueConverter> valueConverters, ValueSetter valueSetter) {
        /*real = workbook;
        if (workbook instanceof AutoWorkbook) {
            int count = 0;
            for (int i = 0; i < writeAdapter.getSheetCount(); i++) {
                count += writeAdapter.getRowCount(i);
            }
            real = AutoWorkbook.getWorkbook(count);
        }*/
        for (PoiWriteListener poiWriteListener : poiWriteListeners) {
            poiWriteListener.onWorkbookCreate(workbook);
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
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetCreate(s, sheet, drawing, workbook);
            }
            int rowCount = writeAdapter.getRowCount(s);
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.createRow(r);
                for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                    poiWriteListener.onRowCreate(r, s, row, sheet, workbook);
                }
                int cellCount = writeAdapter.getCellCount(s, r);
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.createCell(c);
                    for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                        poiWriteListener.onCellCreate(c, r, s, cell, row, sheet, workbook);
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
                    valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, value);
                    for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                        poiWriteListener.onCellValueSet(c, r, s, cell, row, sheet, workbook);
                    }
                }
                for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                    poiWriteListener.onRowValueSet(r, s, row, sheet, workbook);
                }
            }
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetValueSet(s, sheet, drawing, workbook);
            }
        }
        for (PoiWriteListener poiWriteListener : poiWriteListeners) {
            poiWriteListener.onWorkbookValueSet(workbook);
        }
    }

    private static void transferAsync(Executor executor, Workbook workbook, WriteAdapter writeAdapter,
                                      List<PoiWriteListener> poiWriteListeners,
                                      List<ValueConverter> valueConverters, ValueSetter valueSetter) {
        int total = 0;
        for (int s = 0; s < writeAdapter.getSheetCount(); s++) {
            for (int r = 0; r < writeAdapter.getRowCount(s); r++) {
                total += writeAdapter.getCellCount(s, r);
            }
        }
        CountDownLatch countDownLatch = new CountDownLatch(total);
        for (PoiWriteListener poiWriteListener : poiWriteListeners) {
            poiWriteListener.onWorkbookCreate(workbook);
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
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetCreate(s, sheet, drawing, workbook);
            }
            int rowCount = writeAdapter.getRowCount(s);
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.createRow(r);
                for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                    poiWriteListener.onRowCreate(r, s, row, sheet, workbook);
                }
                int cellCount = writeAdapter.getCellCount(s, r);
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.createCell(c);
                    for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                        poiWriteListener.onCellCreate(c, r, s, cell, row, sheet, workbook);
                    }
                    Runnable runnable = new DataTransfer(countDownLatch, writeAdapter, valueConverters,
                            valueSetter, workbook, sheet, row, cell, drawing, s, r, c);
                    executor.execute(runnable);
                    for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                        poiWriteListener.onCellValueSet(c, r, s, cell, row, sheet, workbook);
                    }
                }
                for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                    poiWriteListener.onRowValueSet(r, s, row, sheet, workbook);
                }
            }
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetValueSet(s, sheet, drawing, workbook);
            }
        }
        for (PoiWriteListener poiWriteListener : poiWriteListeners) {
            poiWriteListener.onWorkbookValueSet(workbook);
        }
        try {
            countDownLatch.await();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
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

    static class DataTransfer implements Runnable {

        private CountDownLatch countDownLatch;

        private WriteAdapter writeAdapter;
        private List<ValueConverter> valueConverters;
        private ValueSetter valueSetter;

        private Workbook workbook;
        private Sheet sheet;
        private Row row;
        private Cell cell;
        private Drawing<?> drawing;

        int s;
        int r;
        int c;

        public DataTransfer(CountDownLatch countDownLatch, WriteAdapter writeAdapter, List<ValueConverter> valueConverters,
                            ValueSetter valueSetter, Workbook workbook, Sheet sheet, Row row, Cell cell, Drawing<?> drawing,
                            int s, int r, int c) {
            this.countDownLatch = countDownLatch;
            this.writeAdapter = writeAdapter;
            this.valueConverters = valueConverters;
            this.valueSetter = valueSetter;
            this.workbook = workbook;
            this.sheet = sheet;
            this.row = row;
            this.cell = cell;
            this.drawing = drawing;
            this.s = s;
            this.r = r;
            this.c = c;
        }

        @Override
        public void run() {
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
            valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, value);
            countDownLatch.countDown();
            //System.out.println(countDownLatch.getCount());
        }
    }

    static class WorkbookTransfer implements Runnable {

        Executor executor;
        CountDownLatch countDownLatch;

        Workbook workbook;
        WriteAdapter writeAdapter;
        List<PoiWriteListener> poiWriteListeners;
        List<ValueConverter> valueConverters;
        ValueSetter valueSetter;

        WorkbookTransfer(Executor executor, CountDownLatch countDownLatch, Workbook workbook, WriteAdapter writeAdapter,
                         List<PoiWriteListener> poiWriteListeners,
                         List<ValueConverter> valueConverters, ValueSetter valueSetter) {
            this.executor = executor;
            this.countDownLatch = countDownLatch;

            this.workbook = workbook;
            this.writeAdapter = writeAdapter;
            this.poiWriteListeners = poiWriteListeners;
            this.valueConverters = valueConverters;
            this.valueSetter = valueSetter;
        }

        void transfer() {
            executor.execute(this);
        }

        @Override
        public void run() {
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onWorkbookCreate(workbook);
            }
            int sheetCount = writeAdapter.getSheetCount();
            CountDownLatch cdl = new CountDownLatch(sheetCount);
            for (int s = 0; s < sheetCount; s++) {
                new SheetTransfer(s, executor, cdl, workbook, writeAdapter, poiWriteListeners,
                        valueConverters, valueSetter).transfer();
            }
            try {
                cdl.await();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onWorkbookValueSet(workbook);
            }
            countDownLatch.countDown();
        }
    }

    static class SheetTransfer extends WorkbookTransfer {

        int s;

        SheetTransfer(int s, Executor executor, CountDownLatch countDownLatch, Workbook workbook, WriteAdapter writeAdapter,
                      List<PoiWriteListener> poiWriteListeners,
                      List<ValueConverter> valueConverters, ValueSetter valueSetter) {
            super(executor, countDownLatch, workbook, writeAdapter, poiWriteListeners, valueConverters, valueSetter);
            this.s = s;
        }

        Sheet createSheet(String sheetName) {
            synchronized (SheetTransfer.class) {
                return sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
            }
        }

        @Override
        public void run() {
            String sheetName = writeAdapter.getSheetName(s);
            Sheet sheet = createSheet(sheetName);
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetCreate(s, sheet, drawing, workbook);
            }
            int rowCount = writeAdapter.getRowCount(s);
            CountDownLatch cdl = new CountDownLatch(rowCount);
            for (int r = 0; r < rowCount; r++) {
                new RowTransfer(r, sheet, drawing, s, executor, cdl, workbook, writeAdapter, poiWriteListeners,
                        valueConverters, valueSetter).transfer();
            }
            try {
                cdl.await();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onSheetValueSet(s, sheet, drawing, workbook);
            }
            countDownLatch.countDown();
        }
    }

    static class RowTransfer extends SheetTransfer {

        int r;
        Sheet sheet;
        Drawing<?> drawing;

        RowTransfer(int r, Sheet sheet, Drawing<?> drawing, int s, Executor executor, CountDownLatch countDownLatch,
                    Workbook workbook, WriteAdapter writeAdapter, List<PoiWriteListener> poiWriteListeners,
                    List<ValueConverter> valueConverters, ValueSetter valueSetter) {
            super(s, executor, countDownLatch, workbook, writeAdapter, poiWriteListeners, valueConverters, valueSetter);
            this.r = r;
            this.sheet = sheet;
            this.drawing = drawing;
        }

        Row createRow(int r) {
            synchronized (RowTransfer.class) {
                return sheet.createRow(r);
            }
        }

        @Override
        public void run() {
            Row row = createRow(r);
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onRowCreate(r, s, row, sheet, workbook);
            }
            int cellCount = writeAdapter.getCellCount(s, r);
            CountDownLatch cdl = new CountDownLatch(cellCount);
            for (int c = 0; c < cellCount; c++) {
                new CellTransfer(c, row, r, sheet, drawing, s, executor, cdl, workbook, writeAdapter, poiWriteListeners,
                        valueConverters, valueSetter).transfer();
            }
            try {
                cdl.await();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onRowValueSet(r, s, row, sheet, workbook);
            }
            countDownLatch.countDown();
        }
    }

    static class CellTransfer extends RowTransfer {

        int c;
        Row row;

        CellTransfer(int c, Row row, int r, Sheet sheet, Drawing<?> drawing, int s, Executor executor,
                     CountDownLatch countDownLatch, Workbook workbook,
                     WriteAdapter writeAdapter, List<PoiWriteListener> poiWriteListeners,
                     List<ValueConverter> valueConverters, ValueSetter valueSetter) {
            super(r, sheet, drawing, s, executor, countDownLatch, workbook, writeAdapter, poiWriteListeners, valueConverters, valueSetter);
            this.c = c;
            this.row = row;
        }

        Cell createCell(int c) {
            synchronized (CellTransfer.class) {
                return row.createCell(c);
            }
        }

        @Override
        public void run() {
            Cell cell = createCell(c);
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onCellCreate(c, r, s, cell, row, sheet, workbook);
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
            valueSetter.setValue(s, r, c, cell, row, sheet, drawing, workbook, value);
            for (PoiWriteListener poiWriteListener : poiWriteListeners) {
                poiWriteListener.onCellValueSet(c, r, s, cell, row, sheet, workbook);
            }
            countDownLatch.countDown();
        }
    }
}
