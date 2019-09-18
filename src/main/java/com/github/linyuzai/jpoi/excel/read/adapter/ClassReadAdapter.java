package com.github.linyuzai.jpoi.excel.read.adapter;

public abstract class ClassReadAdapter extends HeaderReadAdapter {
    public static class ReadField {
        private String fieldName;
        private String fieldDescription;
        private int index;

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

        public int getIndex() {
            return index;
        }

        public void setIndex(int index) {
            this.index = index;
        }
    }
}
