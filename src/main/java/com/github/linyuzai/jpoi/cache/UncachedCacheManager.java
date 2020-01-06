package com.github.linyuzai.jpoi.cache;

import com.github.linyuzai.jpoi.excel.JExcelBase;

public class UncachedCacheManager implements CacheManager {

    private static final UncachedCacheManager sInstance = new UncachedCacheManager();

    public static UncachedCacheManager getInstance() {
        return sInstance;
    }

    @Override
    public Object getCache(JExcelBase<?> base, Object source, int s, int r, int c) {
        return null;
    }

    @Override
    public void setCache(JExcelBase<?> base, Object source, Object value, int s, int r, int c) {

    }
}
