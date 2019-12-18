package com.github.linyuzai.jpoi.excel.write.adapter;

public abstract class ClassWriteAdapter extends HeaderWriteAdapter {

    public static class WriteField {
        private String fieldName;
        private String fieldDescription;
        private int width;
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

        public int getWidth() {
            return width;
        }

        public void setWidth(int width) {
            this.width = width;
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

        public WriteField newField(String fieldDescription) {
            WriteField field = new WriteField();
            field.setFieldName(fieldDescription);
            field.setFieldDescription(fieldDescription);
            field.setOrder(order);
            field.setWidth(width);
            field.setAutoSize(autoSize);
            return field;
        }
    }
}
