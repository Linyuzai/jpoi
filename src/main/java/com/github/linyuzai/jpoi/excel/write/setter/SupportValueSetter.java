package com.github.linyuzai.jpoi.excel.write.setter;

import com.github.linyuzai.jpoi.excel.value.combination.CombinationValue;
import com.github.linyuzai.jpoi.excel.value.comment.RichTextStringComment;
import com.github.linyuzai.jpoi.excel.value.comment.StringComment;
import com.github.linyuzai.jpoi.excel.value.comment.SupportComment;
import com.github.linyuzai.jpoi.excel.value.error.SupportErrorValue;
import com.github.linyuzai.jpoi.excel.value.formula.SupportFormula;
import com.github.linyuzai.jpoi.excel.value.picture.*;
import com.github.linyuzai.jpoi.support.SupportValue;
import org.apache.poi.ss.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.Collection;

public class SupportValueSetter extends PoiValueSetter {

    private static SupportValueSetter sInstance = new SupportValueSetter();

    public static SupportValueSetter getInstance() {
        return sInstance;
    }

    @Override
    public void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, Object value) {
        if (value instanceof SupportValue) {
            setSupportValue(s, r, c, cell, row, sheet, drawing, workbook, value);
            return;
        }
        super.setValue(s, r, c, cell, row, sheet, drawing, workbook, value);
    }

    public void setSupportValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, Object value) {
        if (value instanceof SupportPicture) {
            SupportPicture.Location location = ((SupportPicture) value).getLocation();
            SupportPicture.Padding padding = ((SupportPicture) value).getPadding();
            ClientAnchor anchor = drawing.createAnchor(
                    padding.getLeft(), padding.getTop(), padding.getRight(), padding.getBottom(),
                    location.getStartCell(), location.getStartRow(), location.getEndCell(), location.getEndRow());
            anchor.setAnchorType(((SupportPicture) value).getAnchorType());
            createPicture((SupportPicture) value, anchor, workbook, drawing);
        } else if (value instanceof SupportErrorValue) {
            cell.setCellErrorValue(((SupportErrorValue) value).getErrorValue());
        } else if (value instanceof SupportFormula) {
            cell.setCellFormula(((SupportFormula) value).getFormula());
        } else if (value instanceof SupportComment) {
            SupportComment.Location location = ((SupportComment) value).getLocation();
            SupportComment.Padding padding = ((SupportComment) value).getPadding();
            ClientAnchor anchor = drawing.createAnchor(
                    padding.getLeft(), padding.getTop(), padding.getRight(), padding.getBottom(),
                    location.getStartCell(), location.getStartRow(), location.getEndCell(), location.getEndRow());
            createComment(cell, (SupportComment) value, anchor, workbook, drawing);
        } else if (value instanceof CombinationValue) {
            Object combinationValue = ((CombinationValue) value).getValue();
            if (combinationValue instanceof Collection) {
                for (Object o : (Collection) combinationValue) {
                    setValue(s, r, c, cell, row, sheet, drawing, workbook, o);
                }
            } else {
                setValue(s, r, c, cell, row, sheet, drawing, workbook, combinationValue);
            }
        }
    }

    private void createComment(Cell cell, SupportComment value, ClientAnchor anchor, Workbook workbook, Drawing<?> drawing) {
        if (value instanceof StringComment) {
            RichTextString rts = workbook.getCreationHelper().createRichTextString(((StringComment) value).getComment());
            createComment(cell, new RichTextStringComment(rts), anchor, workbook, drawing);
        } else if (value instanceof RichTextStringComment) {
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(((RichTextStringComment) value).getRichTextString());
            cell.setCellComment(comment);
        }
    }

    private void createPicture(SupportPicture value, ClientAnchor anchor, Workbook workbook, Drawing<?> drawing) {
        int type = value.getType();
        if (value instanceof ByteArrayPicture) {
            drawing.createPicture(anchor, workbook.addPicture(((ByteArrayPicture) value).getBytes(), type));
        } else if (value instanceof BufferedImagePicture) {
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            try {
                String format = value.getFormat();
                if (format == null) {
                    if (type == Workbook.PICTURE_TYPE_PNG) {
                        format = "png";
                    } else {
                        format = "jpeg";
                    }
                }
                ImageIO.write(((BufferedImagePicture) value).getBufferedImage(), format, byteArrayOut);
                createPicture(new ByteArrayPicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                        type, value.getFormat(), byteArrayOut.toByteArray()), anchor, workbook, drawing);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else if (value instanceof FilePicture) {
            try {
                File file = ((FilePicture) value).getFile();
                Path path = Paths.get(file.getAbsolutePath());
                String mime = Files.probeContentType(path);
                if (mime.equals("image/png")) {
                    type = Workbook.PICTURE_TYPE_PNG;
                }
                BufferedImage bufferedImage = ImageIO.read(file);
                createPicture(new BufferedImagePicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                        type, value.getFormat(), bufferedImage), anchor, workbook, drawing);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else if (value instanceof Base64Picture) {
            byte[] bytes = Base64.getDecoder().decode(((Base64Picture) value).getBase64());
            createPicture(new ByteArrayPicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                    type, value.getFormat(), bytes), anchor, workbook, drawing);
        }
    }
}
