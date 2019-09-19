package com.github.linyuzai.jpoi.excel.read.adapter;

public class ObjectReadAdapter extends MapReadAdapter {

    private boolean toMap;

    public void setTargetClass(Class<?> cls) {
        ReadField readField = getReadFieldIncludeAnnotation(cls);
        if (readField instanceof AnnotationReadField) {

        }
    }

    public boolean isToMap() {
        return toMap;
    }

    public void setToMap(boolean toMap) {
        this.toMap = toMap;
    }
}
