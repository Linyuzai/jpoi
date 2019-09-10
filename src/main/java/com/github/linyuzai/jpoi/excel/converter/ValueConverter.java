package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.order.Ordered;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public interface ValueConverter extends Ordered {

    Map<String, ValueConverter> cache = new ConcurrentHashMap<>();

    boolean supportValue(int sheet, int row, int cell, Object value);

    Object adaptValue(int sheet, int row, int cell, Object value);
}
