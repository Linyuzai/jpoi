package com.github.linyuzai.jpoi.excel.value;

public interface SupportPicture extends SupportValue {
    PoiPicture.Padding getPadding();

    PoiPicture.Location getLocation();

    int getType();

    String getFormat();

    class Location {
        private int startRow;
        private int startCell;
        private int endRow;
        private int endCell;

        public Location(int startRow, int startCell, int endRow, int endCell) {
            this.startRow = startRow;
            this.startCell = startCell;
            this.endRow = endRow;
            this.endCell = endCell;
        }

        public int getStartRow() {
            return startRow;
        }

        public void setStartRow(int startRow) {
            this.startRow = startRow;
        }

        public int getStartCell() {
            return startCell;
        }

        public void setStartCell(int startCell) {
            this.startCell = startCell;
        }

        public int getEndRow() {
            return endRow;
        }

        public void setEndRow(int endRow) {
            this.endRow = endRow;
        }

        public int getEndCell() {
            return endCell;
        }

        public void setEndCell(int endCell) {
            this.endCell = endCell;
        }
    }

    class Padding {

        private static Padding ZERO = new Padding(0, 0, 0, 0);

        public static Padding getDefault() {
            return ZERO;
        }

        private int top;
        private int bottom;
        private int left;
        private int right;

        public Padding() {
        }

        public Padding(int top, int bottom, int left, int right) {
            this.top = top;
            this.bottom = bottom;
            this.left = left;
            this.right = right;
        }

        public int getTop() {
            return top;
        }

        public void setTop(int top) {
            this.top = top;
        }

        public int getBottom() {
            return bottom;
        }

        public void setBottom(int bottom) {
            this.bottom = bottom;
        }

        public int getLeft() {
            return left;
        }

        public void setLeft(int left) {
            this.left = left;
        }

        public int getRight() {
            return right;
        }

        public void setRight(int right) {
            this.right = right;
        }
    }
}
