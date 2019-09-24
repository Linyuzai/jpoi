package com.github.linyuzai.jpoi.excel.value.picture;

import org.apache.poi.ss.usermodel.ClientAnchor;

import java.awt.image.BufferedImage;

public class BufferedImagePicture extends PoiPicture {

    private BufferedImage bufferedImage;

    public BufferedImagePicture(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
    }

    public BufferedImagePicture(Padding padding, Location location, ClientAnchor.AnchorType anchorType, int type, String format, BufferedImage bufferedImage) {
        super(padding, location, anchorType, type, format);
        this.bufferedImage = bufferedImage;
    }

    public BufferedImage getBufferedImage() {
        return bufferedImage;
    }

    public void setBufferedImage(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
    }
}
