package com.github.linyuzai.jpoi.excel.value.combination;

import com.github.linyuzai.jpoi.support.SupportValue;

public interface SupportCombinationValue extends SupportValue {

    Object getValue();

    void setValue(Object value);

    void addValue(Object value);

    void removeValue(Object value);
}
