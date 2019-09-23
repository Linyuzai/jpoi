package com.github.linyuzai.jpoi.excel.read.sax;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFShape;

import java.util.Iterator;
import java.util.List;

public class SaxDrawing implements Drawing<XSSFShape> {

    private List<XSSFShape> shapes;

    public SaxDrawing(List<XSSFShape> shapes) {
        this.shapes = shapes;
    }

    public List<XSSFShape> getShapes() {
        return shapes;
    }

    public void setShapes(List<XSSFShape> shapes) {
        this.shapes = shapes;
    }

    @Override
    public Picture createPicture(ClientAnchor anchor, int pictureIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Comment createCellComment(ClientAnchor anchor) {
        throw new UnsupportedOperationException();
    }

    @Override
    public ClientAnchor createAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2) {
        throw new UnsupportedOperationException();
    }

    @Override
    public ObjectData createObjectData(ClientAnchor anchor, int storageId, int pictureIndex) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<XSSFShape> iterator() {
        return shapes.iterator();
    }
}
