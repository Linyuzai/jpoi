package com.github.linyuzai.jpoi.excel.cache;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class InMemoryCacheManager implements CacheManager {
    private static final String CACHE_PREFIX = "InMemoryCacheManager@";

    private final Map<Object, Object> cacheMap = new ConcurrentHashMap<>();

    @Override
    public Object getCache(JExcelProcessor<?> processor, Object source, int s, int r, int c) {
        return cacheMap.get(getCachePrefix() + source);
    }

    @Override
    public void setCache(JExcelProcessor<?> processor, Object source, Object value, int s, int r, int c) {
        cacheMap.put(getCachePrefix() + source, value);
    }

    public String getCachePrefix() {
        return CACHE_PREFIX;
    }
}
