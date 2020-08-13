package com.github.linyuzai.jpoi.excel.post;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.handler.ExceptionValue;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;

import java.util.Collection;
import java.util.List;

public interface PostProcessor {

    List<ExceptionValue> processPost(Collection<? extends PostValue> postValues, JExcelProcessor<?> context);
}
