package com.github.linyuzai.jpoi.excel.write.style;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class JCellStyle {
    private VerticalAlignment verticalAlignment;

    private HorizontalAlignment horizontalAlignment;

    private JCellBorder border;

    private short dataFormat;

    private JCellFill fill;

    private JCellFont font;

    private boolean hidden;

    private short indention;

    private boolean locked;

    private boolean quotePrefixed;

    private short rotation;

    private boolean shrinkToFit;

    private boolean wrapText;

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public JCellBorder getBorder() {
        return border;
    }

    public void setBorder(JCellBorder border) {
        this.border = border;
    }

    public short getDataFormat() {
        return dataFormat;
    }

    public void setDataFormat(short dataFormat) {
        this.dataFormat = dataFormat;
    }

    public JCellFill getFill() {
        return fill;
    }

    public void setFill(JCellFill fill) {
        this.fill = fill;
    }

    public JCellFont getFont() {
        return font;
    }

    public void setFont(JCellFont font) {
        this.font = font;
    }

    public boolean isHidden() {
        return hidden;
    }

    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    public short getIndention() {
        return indention;
    }

    public void setIndention(short indention) {
        this.indention = indention;
    }

    public boolean isLocked() {
        return locked;
    }

    public void setLocked(boolean locked) {
        this.locked = locked;
    }

    public boolean isQuotePrefixed() {
        return quotePrefixed;
    }

    public void setQuotePrefixed(boolean quotePrefixed) {
        this.quotePrefixed = quotePrefixed;
    }

    public short getRotation() {
        return rotation;
    }

    public void setRotation(short rotation) {
        this.rotation = rotation;
    }

    public boolean isShrinkToFit() {
        return shrinkToFit;
    }

    public void setShrinkToFit(boolean shrinkToFit) {
        this.shrinkToFit = shrinkToFit;
    }

    public boolean isWrapText() {
        return wrapText;
    }

    public void setWrapText(boolean wrapText) {
        this.wrapText = wrapText;
    }
}
