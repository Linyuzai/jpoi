package com.github.linyuzai.jpoi.excel.cache;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class InMemoryUrlCacheManager extends InMemoryCacheManager {
    private static final String CACHE_PREFIX = "InMemoryUrlCacheManager@";

    private final Map<Object, Object> cacheMap = new ConcurrentHashMap<>();

    @Override
    public Object getCache(JExcelProcessor<?> context, Object source, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            return super.getCache(context, ((UrlPostValueHolder) source).getSource(), s, r, c);
        }
        if (source instanceof String) {
            return super.getCache(context, source, s, r, c);
        }
        return null;
    }

    @Override
    public void setCache(JExcelProcessor<?> context, Object source, Object value, int s, int r, int c) {
        if (source instanceof UrlPostValueHolder) {
            super.setCache(context, ((UrlPostValueHolder) source).getSource(), value, s, r, c);
        }
        if (source instanceof String) {
            super.setCache(context, source, value, s, r, c);
        }
    }

    @Override
    public String getCachePrefix() {
        return CACHE_PREFIX;
    }
}
