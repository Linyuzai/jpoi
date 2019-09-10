package com.github.linyuzai.jpoi.excel.value;

public class PoiPicture implements SupportPicture {

    private Padding padding;
    private Location location;
    private int type;
    private String format;

    public PoiPicture() {
    }

    public PoiPicture(Padding padding, Location location, int type, String format) {
        this.padding = padding;
        this.location = location;
        this.type = type;
        this.format = format;
    }

    @Override
    public Padding getPadding() {
        return padding;
    }

    public void setPadding(Padding padding) {
        this.padding = padding;
    }

    @Override
    public Location getLocation() {
        return location;
    }

    public void setLocation(Location location) {
        this.location = location;
    }

    @Override
    public int getType() {
        return type;
    }

    public void setType(int type) {
        this.type = type;
    }

    @Override
    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }
}
