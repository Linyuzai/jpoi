package com.github.linyuzai.jpoi.excel.cache;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;

public interface CacheManager {

    Object getCache(JExcelProcessor<?> processor, Object source, int s, int r, int c);

    void setCache(JExcelProcessor<?> processor, Object source, Object value, int s, int r, int c);
}
