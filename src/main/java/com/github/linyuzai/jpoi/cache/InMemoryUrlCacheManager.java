package com.github.linyuzai.jpoi.cache;

import com.github.linyuzai.jpoi.excel.JExcelBase;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class InMemoryUrlCacheManager implements CacheManager {
    private static final String CACHE_PREFIX = "UrlPostValueHolder@";

    private final Map<Object, Object> cacheMap = new ConcurrentHashMap<>();

    @Override
    public Object getCache(JExcelBase<?> base, Object source, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            return cacheMap.get(CACHE_PREFIX + ((UrlPostValueHolder) source).getSource());
        }
        return null;
    }

    @Override
    public void setCache(JExcelBase<?> base, Object source, Object value, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            cacheMap.put(CACHE_PREFIX + ((UrlPostValueHolder) source).getSource(), value);
        }
    }
}
