package com.github.linyuzai.jpoi.excel.cache;

import com.github.linyuzai.jpoi.excel.JExcelProcessor;

public class UncachedCacheManager implements CacheManager {

    private static final UncachedCacheManager sInstance = new UncachedCacheManager();

    public static UncachedCacheManager getInstance() {
        return sInstance;
    }

    @Override
    public Object getCache(JExcelProcessor<?> context, Object source, int s, int r, int c) {
        return null;
    }

    @Override
    public void setCache(JExcelProcessor<?> context, Object source, Object value, int s, int r, int c) {

    }
}
