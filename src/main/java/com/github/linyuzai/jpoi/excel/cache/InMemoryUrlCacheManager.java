package com.github.linyuzai.jpoi.excel.cache;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class InMemoryUrlCacheManager extends InMemoryCacheManager {
    private static final String CACHE_PREFIX = "InMemoryUrlCacheManager@";

    private final Map<Object, Object> cacheMap = new ConcurrentHashMap<>();

    @Override
    public Object getCache(JExcelProcessor<?> processor, Object source, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            return super.getCache(processor, ((UrlPostValueHolder) source).getSource(), s, r, c);
        }
        if (source instanceof String) {
            return super.getCache(processor, source, s, r, c);
        }
        return null;
    }

    @Override
    public void setCache(JExcelProcessor<?> processor, Object source, Object value, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            super.setCache(processor, ((UrlPostValueHolder) source).getSource(), value, s, r, c);
        }
        if (source instanceof String) {
            super.setCache(processor, source, value, s, r, c);
        }
    }

    @Override
    public String getCachePrefix() {
        return CACHE_PREFIX;
    }
}
