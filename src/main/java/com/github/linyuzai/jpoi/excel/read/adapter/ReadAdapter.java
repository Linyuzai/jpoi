package com.github.linyuzai.jpoi.excel.read.adapter;

public interface ReadAdapter {

    void readCell(Object value, int s, int r, int c, int sCount, int rCount, int cCount) throws Throwable;

    Object getValue();
}
