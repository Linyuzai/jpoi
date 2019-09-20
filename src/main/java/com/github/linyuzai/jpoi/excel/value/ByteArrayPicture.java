package com.github.linyuzai.jpoi.excel.value;

import org.apache.poi.ss.usermodel.ClientAnchor;

public class ByteArrayPicture extends PoiPicture {

    private byte[] bytes;

    public ByteArrayPicture(byte[] bytes) {
        this.bytes = bytes;
    }

    public ByteArrayPicture(Padding padding, Location location, ClientAnchor.AnchorType anchorType, int type, String format, byte[] bytes) {
        super(padding, location, anchorType, type, format);
        this.bytes = bytes;
    }

    public byte[] getBytes() {
        return bytes;
    }

    public void setBytes(byte[] bytes) {
        this.bytes = bytes;
    }
}
