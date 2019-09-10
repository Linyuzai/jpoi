package com.github.linyuzai.jpoi.excel.value;

public class ByteArrayPicture extends PoiPicture {

    private byte[] bytes;

    public ByteArrayPicture(byte[] bytes) {
        this.bytes = bytes;
    }

    public ByteArrayPicture(Padding padding, Location location, int type, String format, byte[] bytes) {
        super(padding, location, type, format);
        this.bytes = bytes;
    }

    public byte[] getBytes() {
        return bytes;
    }

    public void setBytes(byte[] bytes) {
        this.bytes = bytes;
    }
}
