package com.github.linyuzai.jpoi.excel.processor;

import com.github.linyuzai.jpoi.excel.handler.ExcelExceptionHandler;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;

import java.util.Collection;
import java.util.Collections;
import java.util.List;

public class EmptyPostProcessor implements PostProcessor {

    private static final EmptyPostProcessor sInstance = new EmptyPostProcessor();

    public static EmptyPostProcessor getInstance() {
        return sInstance;
    }

    @Override
    public List<Throwable> processPost(Collection<? extends PostValue> postValues, ExcelExceptionHandler handler) {
        return Collections.emptyList();
    }
}
