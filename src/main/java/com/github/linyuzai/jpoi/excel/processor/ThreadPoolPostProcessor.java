package com.github.linyuzai.jpoi.excel.processor;

import com.github.linyuzai.jpoi.excel.value.post.PostValue;

import java.util.Collection;
import java.util.List;
import java.util.Vector;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;

public abstract class ThreadPoolPostProcessor implements PostProcessor {

    private ExecutorService executorService;

    public ThreadPoolPostProcessor(ExecutorService executorService) {
        this.executorService = executorService;
    }

    @Override
    public List<Throwable> processPost(Collection<? extends PostValue> postValues) {
        List<Throwable> throwableRecords = new Vector<>();
        CountDownLatch cdl = new CountDownLatch(getCountDownLatchCount(postValues));
        for (PostValue postValue : postValues) {
            executorService.execute(() -> {
                try {
                    processPostValue(postValue);
                } catch (Throwable e) {
                    throwableRecords.add(e);
                } finally {
                    cdl.countDown();
                }
            });
        }
        try {
            cdl.await();
        } catch (InterruptedException e) {
            throwableRecords.add(e);
        }
        return throwableRecords;
    }

    public abstract void processPostValue(PostValue postValue) throws Throwable;

    public int getCountDownLatchCount(Collection<? extends PostValue> postValues) {
        return postValues.size();
    }

    public ExecutorService getExecutorService() {
        return executorService;
    }

    public void setExecutorService(ExecutorService executorService) {
        this.executorService = executorService;
    }
}
