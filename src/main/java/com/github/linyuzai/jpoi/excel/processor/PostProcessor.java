package com.github.linyuzai.jpoi.excel.processor;

import com.github.linyuzai.jpoi.excel.handler.ExcelExceptionHandler;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;

import java.util.Collection;
import java.util.List;

public interface PostProcessor {

    List<Throwable> processPost(Collection<? extends PostValue> postValues, ExcelExceptionHandler handler);
}
