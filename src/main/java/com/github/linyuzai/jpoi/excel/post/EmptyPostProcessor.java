package com.github.linyuzai.jpoi.excel.post;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.handler.ExceptionValue;
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
    public List<ExceptionValue> processPost(Collection<? extends PostValue> postValues, JExcelProcessor<?> context) {
        return Collections.emptyList();
    }
}
