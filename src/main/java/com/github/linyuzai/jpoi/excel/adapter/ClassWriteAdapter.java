package com.github.linyuzai.jpoi.excel.adapter;

public abstract class ClassWriteAdapter extends HeaderWriteAdapter {

    public static class FieldData {
        private String fieldName;
        private String fieldDescription;

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
    }
}
