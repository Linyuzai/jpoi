package com.github.linyuzai.jpoi.excel.value.picture;

import org.apache.poi.ss.usermodel.ClientAnchor;

public class Base64Picture extends PoiPicture {

    private String base64;

    public Base64Picture(String base64) {
        this.base64 = base64;
    }

    public Base64Picture(Padding padding, Location location, ClientAnchor.AnchorType anchorType, int type, String format, String base64) {
        super(padding, location, anchorType, type, format);
        this.base64 = base64;
    }

    public String getBase64() {
        return base64;
    }

    public void setBase64(String base64) {
        this.base64 = base64;
    }
}
