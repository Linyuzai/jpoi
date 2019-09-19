package com.github.linyuzai.jpoi.excel.read.adapter;

public abstract class ClassReadAdapter<T> extends ListReadAdapter<T> {
    public static class ReadField {
        private int index;
        private String fieldName;
        private String fieldDescription;
        private int order;

        public int getIndex() {
            return index;
        }

        public void setIndex(int index) {
            this.index = index;
        }

        public String getFieldName() {
            return fieldName;
        }

        public void setFieldName(String fieldName) {
            this.fieldName = fieldName;
        }

        public String getFieldDescription() {
            return fieldDescription;
        }

        public void setFieldDescription(String fieldDescription) {
            this.fieldDescription = fieldDescription;
        }

        public int getOrder() {
            return order;
        }

        public void setOrder(int order) {
            this.order = order;
        }
    }
}
