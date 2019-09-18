package com.github.linyuzai.jpoi.excel.write.value;

import org.apache.poi.ss.usermodel.ClientAnchor;

import java.io.File;

public class FilePicture extends PoiPicture {

    private File file;

    public FilePicture(File file) {
        this.file = file;
    }

    public FilePicture(Padding padding, Location location, ClientAnchor.AnchorType anchorType, int type, String format, File file) {
        super(padding, location, anchorType, type, format);
        this.file = file;
    }

    public FilePicture(String path) {
        this.file = new File(path);
    }

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }

    public void setFile(String path) {
        this.file = new File(path);
    }
}
