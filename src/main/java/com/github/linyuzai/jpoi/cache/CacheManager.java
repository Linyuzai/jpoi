package com.github.linyuzai.jpoi.cache;

import com.github.linyuzai.jpoi.excel.JExcelBase;

public interface CacheManager {

    Object getCache(JExcelBase<?> base, Object source, int s, int r, int c);

    void setCache(JExcelBase<?> base, Object source, Object value, int s, int r, int c);
}
