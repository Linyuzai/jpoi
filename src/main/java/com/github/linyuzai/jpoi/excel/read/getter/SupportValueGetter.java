package com.github.linyuzai.jpoi.excel.read.getter;

import com.github.linyuzai.jpoi.excel.value.ByteArrayPicture;
import org.apache.poi.ss.usermodel.*;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class SupportValueGetter extends PoiValueGetter {

    public static SupportValueGetter sInstance = new SupportValueGetter();

    public static SupportValueGetter getInstance() {
        return sInstance;
    }

    private Map<Integer, Map<Integer, ByteArrayPicture>> pictureMap = null;

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook) {
        if (pictureMap == null) {
            synchronized (this) {
                if (pictureMap == null) {
                    Map<Integer, Map<Integer, ByteArrayPicture>> pictures = new HashMap<>();
                    if (drawing != null) {
                        for (Shape shape : drawing) {
                            if (shape instanceof Picture) {
                                ByteArrayPicture bap = new ByteArrayPicture(((Picture) shape).getPictureData().getData());
                                bap.setFormat(((Picture) shape).getPictureData().getMimeType());
                                bap.setType(((Picture) shape).getPictureData().getPictureType());
                                ClientAnchor anchor = ((Picture) shape).getClientAnchor();
                                int rStart = anchor.getRow1();
                                int cStart = anchor.getCol1();
                                Map<Integer, ByteArrayPicture> pictureCell = pictures.computeIfAbsent(rStart, HashMap::new);
                                pictureCell.put(cStart, bap);
                            }
                        }
                    }
                    pictureMap = Collections.unmodifiableMap(pictures);
                }
            }
        }
        Map<Integer, ByteArrayPicture> fromRow = pictureMap.get(r);
        if (fromRow != null) {
            ByteArrayPicture fromCell = fromRow.get(c);
            if (fromCell != null) {
                return fromCell;
            }
        }
        return super.getValue(s, r, c, cell, row, sheet, drawing, workbook);
    }
}
