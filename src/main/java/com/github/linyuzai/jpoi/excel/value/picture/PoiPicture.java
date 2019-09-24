package com.github.linyuzai.jpoi.excel.value.picture;

import com.github.linyuzai.jpoi.excel.value.ClientAnchorValue;
import org.apache.poi.ss.usermodel.ClientAnchor;

public class PoiPicture extends ClientAnchorValue implements SupportPicture {
    private int type;
    private String format;

    public PoiPicture() {
    }

    public PoiPicture(Padding padding, Location location, ClientAnchor.AnchorType anchorType, int type, String format) {
        super(padding, location, anchorType);
        this.type = type;
        this.format = format;
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
