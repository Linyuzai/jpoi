package com.github.linyuzai.jpoi.excel.write.style;

import org.apache.poi.ss.usermodel.FillPatternType;

public class JCellFill {
    private FillPatternType pattern;

    private short foregroundColor;

    private short backgroundColor;

    public FillPatternType getPattern() {
        return pattern;
    }

    public void setPattern(FillPatternType pattern) {
        this.pattern = pattern;
    }

    public short getForegroundColor() {
        return foregroundColor;
    }

    public void setForegroundColor(short foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    public short getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }
}
