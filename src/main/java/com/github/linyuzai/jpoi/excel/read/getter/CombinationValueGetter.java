package com.github.linyuzai.jpoi.excel.read.getter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.excel.value.combination.ListCombinationValue;
import com.github.linyuzai.jpoi.excel.value.comment.StringComment;
import com.github.linyuzai.jpoi.excel.value.comment.SupportComment;
import com.github.linyuzai.jpoi.excel.value.picture.ByteArrayPicture;
import com.github.linyuzai.jpoi.excel.value.picture.SupportPicture;
import org.apache.poi.ss.usermodel.*;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class CombinationValueGetter extends SupportValueGetter {

    public static CombinationValueGetter sInstance = new CombinationValueGetter();

    public static CombinationValueGetter getInstance() {
        return sInstance;
    }

    private Map<Integer, Map<Integer, ByteArrayPicture>> pictureMap = null;

    @Override
    public Object getValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper) {
        if (pictureMap == null) {
            synchronized (this) {
                if (pictureMap == null) {
                    Map<Integer, Map<Integer, ByteArrayPicture>> pictures = new HashMap<>();
                    if (drawing != null) {
                        for (Shape shape : drawing) {
                            if (shape instanceof Picture) {
                                PictureData pictureData = ((Picture) shape).getPictureData();
                                if (pictureData == null) {
                                    continue;
                                }
                                ByteArrayPicture bap = new ByteArrayPicture(pictureData.getData());
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
        SupportPicture picture = null;
        if (fromRow != null) {
            picture = fromRow.get(c);
        }
        Comment cellComment = cell.getCellComment();
        SupportComment comment = null;
        if (cellComment != null) {
            comment = new StringComment(cellComment.getString().getString());
        }
        Object cellData = super.getValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper);
        if (picture != null || comment != null) {
            CombinationValue combinationValue = new ListCombinationValue();
            combinationValue.addValue(cellData);
            if (comment != null) {
                combinationValue.addValue(comment);
            }
            if (picture != null) {
                combinationValue.addValue(picture);
            }
            return combinationValue;
        } else {
            return cellData;
        }
    }
}
