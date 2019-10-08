package com.github.linyuzai.jpoi.excel.write.style;

public class JRowStyle {
    private short height;

    private float heightInPoints;

    private boolean zeroHeight;

    private JCellStyle cellStyle;

    public short getHeight() {
        return height;
    }

    public void setHeight(short height) {
        this.height = height;
    }

    public float getHeightInPoints() {
        return heightInPoints;
    }

    public void setHeightInPoints(float heightInPoints) {
        this.heightInPoints = heightInPoints;
    }

    public boolean isZeroHeight() {
        return zeroHeight;
    }

    public void setZeroHeight(boolean zeroHeight) {
        this.zeroHeight = zeroHeight;
    }

    public JCellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(JCellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}
