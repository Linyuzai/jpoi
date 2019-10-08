package com.github.linyuzai.jpoi.excel.write.style;

public class JCellFont {
    private boolean bold;

    private int charSet;

    private short color;

    private short fontHeight;

    private short fontHeightInPoints;

    private String fontName;

    private boolean italic;

    private boolean strikeout;

    private short typeOffset;

    private byte underline;

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public int getCharSet() {
        return charSet;
    }

    public void setCharSet(int charSet) {
        this.charSet = charSet;
    }

    public short getColor() {
        return color;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public short getFontHeight() {
        return fontHeight;
    }

    public void setFontHeight(short fontHeight) {
        this.fontHeight = fontHeight;
    }

    public short getFontHeightInPoints() {
        return fontHeightInPoints;
    }

    public void setFontHeightInPoints(short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public boolean isItalic() {
        return italic;
    }

    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    public boolean isStrikeout() {
        return strikeout;
    }

    public void setStrikeout(boolean strikeout) {
        this.strikeout = strikeout;
    }

    public short getTypeOffset() {
        return typeOffset;
    }

    public void setTypeOffset(short typeOffset) {
        this.typeOffset = typeOffset;
    }

    public byte getUnderline() {
        return underline;
    }

    public void setUnderline(byte underline) {
        this.underline = underline;
    }
}
