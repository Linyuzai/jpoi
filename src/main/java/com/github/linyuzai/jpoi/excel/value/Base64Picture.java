package com.github.linyuzai.jpoi.excel.value;

public class Base64Picture extends PoiPicture {

    private String base64;

    public Base64Picture(String base64) {
        this.base64 = base64;
    }

    public String getBase64() {
        return base64;
    }

    public void setBase64(String base64) {
        this.base64 = base64;
    }
}
