package com.github.linyuzai.jpoi.excel.read.adapter;

public class ObjectReadAdapter extends MapReadAdapter {

    private Class<?> cls;
    private boolean toMap;

    public void setTargetClass(Class<?> cls) {
        this.cls = cls;
        ReadField readField = getReadFieldIncludeAnnotation(cls);
        if (readField instanceof AnnotationReadField) {
            setToMap(((AnnotationReadField) readField).isToMap());
            setAnnotationOnly(((AnnotationReadField) readField).isAnnotationOnly());
        }
    }

    public boolean isToMap() {
        return toMap;
    }

    public void setToMap(boolean toMap) {
        this.toMap = toMap;
    }

    @Override
    public Object createContainer(Object value, int s, int r, int c, int sCount, int rCount, int cCount) {
        try {
            return cls.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void adaptValue(Object cellContainer, Object value, int s, int r, int c, int sCount, int rCount, int cCount) {

    }
}
