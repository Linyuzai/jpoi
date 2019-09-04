package com.github.linyuzai.jpoi.excel.converter;

import com.github.linyuzai.jpoi.order.Ordered;

public interface ValueConverter extends Ordered {

    boolean supportValue(int sheet, int row, int cell, Object value);

    Object adaptValue(int sheet, int row, int cell, Object value);
}
