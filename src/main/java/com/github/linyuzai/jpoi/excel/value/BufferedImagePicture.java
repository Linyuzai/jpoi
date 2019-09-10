package com.github.linyuzai.jpoi.excel.value;

import java.awt.image.BufferedImage;

public class BufferedImagePicture extends PoiPicture {

    private BufferedImage bufferedImage;

    public BufferedImagePicture(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
    }

    public BufferedImagePicture(Padding padding, Location location, int type, String format, BufferedImage bufferedImage) {
        super(padding, location, type, format);
        this.bufferedImage = bufferedImage;
    }

    public BufferedImage getBufferedImage() {
        return bufferedImage;
    }

    public void setBufferedImage(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
    }
}
