package com.github.linyuzai.jpoi.excel.value;

import java.awt.image.BufferedImage;

public class BufferedImagePicture extends PoiPicture {

    private BufferedImage bufferedImage;

    public BufferedImagePicture(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
    }

    public BufferedImage getBufferedImage() {
        return bufferedImage;
    }

    public void setBufferedImage(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
    }
}
