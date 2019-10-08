package com.github.linyuzai.jpoi.excel.write.style;

import org.apache.poi.ss.usermodel.BorderStyle;

public class JCellBorder {
    private BorderStyle top;

    private short topColor;

    private BorderStyle right;

    private short rightColor;

    private BorderStyle bottom;

    private short bottomColor;

    private BorderStyle left;

    private short leftColor;

    public BorderStyle getTop() {
        return top;
    }

    public void setTop(BorderStyle top) {
        this.top = top;
    }

    public short getTopColor() {
        return topColor;
    }

    public void setTopColor(short topColor) {
        this.topColor = topColor;
    }

    public BorderStyle getRight() {
        return right;
    }

    public void setRight(BorderStyle right) {
        this.right = right;
    }

    public short getRightColor() {
        return rightColor;
    }

    public void setRightColor(short rightColor) {
        this.rightColor = rightColor;
    }

    public BorderStyle getBottom() {
        return bottom;
    }

    public void setBottom(BorderStyle bottom) {
        this.bottom = bottom;
    }

    public short getBottomColor() {
        return bottomColor;
    }

    public void setBottomColor(short bottomColor) {
        this.bottomColor = bottomColor;
    }

    public BorderStyle getLeft() {
        return left;
    }

    public void setLeft(BorderStyle left) {
        this.left = left;
    }

    public short getLeftColor() {
        return leftColor;
    }

    public void setLeftColor(short leftColor) {
        this.leftColor = leftColor;
    }
}
