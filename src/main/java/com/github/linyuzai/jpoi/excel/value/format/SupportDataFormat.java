package com.github.linyuzai.jpoi.excel.value.format;

import com.github.linyuzai.jpoi.support.SupportValue;

public interface SupportDataFormat extends SupportValue {

    double getValue();

    String getFormatValue();

    short getDataFormat();
}
