package com.github.linyuzai.jpoi.excel;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;

import java.util.List;

public class JExcelBase<T> {

    public static Object convertValue(List<ValueConverter> valueConverters, int s, int r, int c, Object o) {
        ValueConverter valueConverter = null;
        for (ValueConverter vc : valueConverters) {
            if (vc.supportValue(s, r, c, o)) {
                valueConverter = vc;
                break;
            }
        }
        if (valueConverter == null) {
            throw new RuntimeException("No value converter matched");
        }
        return valueConverter.convertValue(s, r, c, o);
    }
}
