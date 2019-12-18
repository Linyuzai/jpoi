package com.github.linyuzai.jpoi.excel.read;

import java.util.Collection;
import java.util.Iterator;
import java.util.List;

public class JExcelReader {
    private JExcelAnalyzer.Values values;

    public JExcelReader(JExcelAnalyzer.Values values) {
        this.values = values;
    }

    @SuppressWarnings("unchecked")
    public <T> T getValue() {
        return (T) values.getValue();
    }

    public Object first() {
        return index(0);
    }

    public JExcelReader firstReader() {
        return new JExcelReader(new JExcelAnalyzer.Values(first(), values.getThrowableRecords()));
    }

    public Object last() {
        Object value = values.getValue();
        if (value instanceof Collection) {
            return index(((Collection) value).size() - 1);
        } else {
            return value;
        }
    }

    public JExcelReader lastReader() {
        return new JExcelReader(new JExcelAnalyzer.Values(last(), values.getThrowableRecords()));
    }

    public Object index(int index) {
        Object value = values.getValue();
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
        return new JExcelReader(new JExcelAnalyzer.Values(index(index), values.getThrowableRecords()));
    }

    public List<Throwable> getThrowableRecords() {
        return values.getThrowableRecords();
    }
}
