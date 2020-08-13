package com.github.linyuzai.jpoi.excel.post;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.handler.ExceptionValue;
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
    public List<ExceptionValue> processPost(Collection<? extends PostValue> postValues, JExcelProcessor<?> context) {
        List<ExceptionValue> exceptionValues = new Vector<>();
        final Collection<? extends PostValue> processValues = getProcessPostValues(postValues);
        CountDownLatch cdl = new CountDownLatch(processValues.size());
        for (PostValue postValue : processValues) {
            executorService.execute(() -> {
                try {
                    processPostValue(postValue);
                } catch (Throwable e) {
                    context.getExceptionHandler().handle(context,
                            postValue.getSheetIndex(),
                            postValue.getRowIndex(),
                            postValue.getCellIndex(),
                            postValue.getCell(),
                            postValue.getRow(),
                            postValue.getSheet(),
                            postValue.getWorkbook(),
                            e);
                    exceptionValues.add(new ExceptionValue(
                            postValue.getSheetIndex(),
                            postValue.getRowIndex(),
                            postValue.getCellIndex(),
                            e));
                } finally {
                    cdl.countDown();
                }
            });
        }
        try {
            cdl.await();
        } catch (InterruptedException e) {
            context.getExceptionHandler().handle(context,
                    -1,
                    -1,
                    -1,
                    null,
                    null,
                    null,
                    null,
                    e);
            exceptionValues.add(new ExceptionValue(
                    -1,
                    -1,
                    -1,
                    e));
        }
        return exceptionValues;
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
