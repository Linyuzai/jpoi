package com.github.linyuzai.jpoi.excel.read;

import java.util.Collection;
import java.util.Iterator;

public class JExcelReader {
    private Object value;

    public JExcelReader(Object value) {
        this.value = value;
    }

    public Object getValue() {
        return value;
    }

    public Object first() {
        return index(0);
    }

    public JExcelReader firstReader() {
        return new JExcelReader(first());
    }

    public Object last() {
        if (value instanceof Collection) {
            return index(((Collection) value).size() - 1);
        } else {
            return value;
        }
    }

    public JExcelReader lastReader() {
        return new JExcelReader(last());
    }

    public Object index(int index) {
        if (value instanceof Collection) {
            Iterator iterator = ((Collection) value).iterator();
            for (int i = 0; iterator.hasNext(); i++) {
                Object o = iterator.next();
                if (i == index) {
                    return o;
                }
            }
            return null;
        } else {
            return value;
        }
    }

    public JExcelReader indexReader(int index) {
        return new JExcelReader(index(index));
    }
}
