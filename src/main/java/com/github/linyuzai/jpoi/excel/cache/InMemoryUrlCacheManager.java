package com.github.linyuzai.jpoi.excel.cache;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class InMemoryUrlCacheManager implements CacheManager {
    private static final String CACHE_PREFIX = "UrlPostValueHolder@";

    private final Map<Object, Object> cacheMap = new ConcurrentHashMap<>();

    @Override
    public Object getCache(JExcelProcessor<?> processor, Object source, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            return cacheMap.get(CACHE_PREFIX + ((UrlPostValueHolder) source).getSource());
        }
        return null;
    }

    @Override
    public void setCache(JExcelProcessor<?> processor, Object source, Object value, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            cacheMap.put(CACHE_PREFIX + ((UrlPostValueHolder) source).getSource(), value);
        }
    }
}
