package com.github.linyuzai.jpoi.excel.adapter;

public abstract class ClassWriteAdapter extends HeaderWriteAdapter {

    public static class FieldData {
        private String fieldName;
        private String fieldDescription;
        private boolean autoSize;
        private int order;

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

        public boolean isAutoSize() {
            return autoSize;
        }

        public void setAutoSize(boolean autoSize) {
            this.autoSize = autoSize;
        }

        public int getOrder() {
            return order;
        }

        public void setOrder(int order) {
            this.order = order;
        }
    }
}
