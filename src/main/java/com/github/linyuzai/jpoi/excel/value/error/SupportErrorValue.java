package com.github.linyuzai.jpoi.excel.value.error;

import com.github.linyuzai.jpoi.support.SupportValue;

public interface SupportErrorValue extends SupportValue {

    byte getErrorValue();

    void setErrorValue(byte value);
}
