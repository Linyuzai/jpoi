package com.github.linyuzai.jpoi.excel.value.combination;

public interface CombinationValue {

    Object getValue();

    void setValue(Object value);

    void addValue(Object value);

    void removeValue(Object value);
}
