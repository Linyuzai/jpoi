package com.github.linyuzai.jpoi.excel.processor;

import com.github.linyuzai.jpoi.excel.handler.ExcelExceptionHandler;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;

import java.util.Collection;
import java.util.List;
import java.util.Vector;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public abstract class ThreadPoolPostProcessor implements PostProcessor {

    private ExecutorService executorService;

    public ThreadPoolPostProcessor() {
        this(Executors.newCachedThreadPool());
    }

    public ThreadPoolPostProcessor(ExecutorService executorService) {
        this.executorService = executorService;
    }

    @Override
    public List<Throwable> processPost(Collection<? extends PostValue> postValues, ExcelExceptionHandler handler) {
        List<Throwable> throwableRecords = new Vector<>();
        final Collection<? extends PostValue> processValues = getProcessPostValues(postValues);
        CountDownLatch cdl = new CountDownLatch(processValues.size());
        for (PostValue postValue : processValues) {
            executorService.execute(() -> {
                try {
                    processPostValue(postValue);
                } catch (Throwable e) {
                    handler.handle(postValue.getSheetIndex(), postValue.getRowIndex(), postValue.getCellIndex(),
                            postValue.getCell(), postValue.getRow(), postValue.getSheet(), postValue.getWorkbook(), e);
                    throwableRecords.add(e);
                } finally {
                    cdl.countDown();
                }
            });
        }
        try {
            cdl.await();
        } catch (InterruptedException e) {
            handler.handle(-1, -1, -1, null, null, null, null, e);
            throwableRecords.add(e);
        }
        return throwableRecords;
    }

    public abstract void processPostValue(PostValue postValue) throws Throwable;

    public Collection<? extends PostValue> getProcessPostValues(Collection<? extends PostValue> postValues) {
        return postValues;
    }

    public ExecutorService getExecutorService() {
        return executorService;
    }

    public void setExecutorService(ExecutorService executorService) {
        this.executorService = executorService;
    }
}
