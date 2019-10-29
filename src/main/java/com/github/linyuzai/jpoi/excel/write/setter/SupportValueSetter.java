package com.github.linyuzai.jpoi.excel.write.setter;

import com.github.linyuzai.jpoi.excel.value.comment.RichTextStringComment;
import com.github.linyuzai.jpoi.excel.value.comment.StringComment;
import com.github.linyuzai.jpoi.excel.value.comment.SupportComment;
import com.github.linyuzai.jpoi.excel.value.error.SupportErrorValue;
import com.github.linyuzai.jpoi.excel.value.formula.SupportFormula;
import com.github.linyuzai.jpoi.excel.value.picture.*;
import com.github.linyuzai.jpoi.support.SupportValue;
import org.apache.poi.ss.usermodel.*;

import javax.imageio.ImageIO;
import java.io.*;
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
    public void setValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper, Object value) {
        if (value instanceof SupportValue) {
            setSupportValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
            return;
        }
        super.setValue(s, r, c, cell, row, sheet, drawing, workbook, creationHelper, value);
    }

    public void setSupportValue(int s, int r, int c, Cell cell, Row row, Sheet sheet, Drawing<?> drawing, Workbook workbook, CreationHelper creationHelper, Object value) {
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
        }
    }

    private void createComment(Cell cell, SupportComment value, ClientAnchor anchor, Workbook workbook, Drawing<?> drawing) {
        if (value instanceof StringComment) {
            RichTextString rts = workbook.getCreationHelper().createRichTextString(value.getComment());
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
            if (type < Workbook.PICTURE_TYPE_EMF || type > Workbook.PICTURE_TYPE_DIB) {
                throw new IllegalArgumentException("Picture type[" + type + "] is unsupported");
            }
            drawing.createPicture(anchor, workbook.addPicture(((ByteArrayPicture) value).getBytes(), type));
        } else if (value instanceof BufferedImagePicture) {
            try {
                String format = value.getFormat();
                if (format == null) {
                    switch (type) {
                        /*case Workbook.PICTURE_TYPE_EMF:
                            format = "emf";
                            break;
                        case Workbook.PICTURE_TYPE_WMF:
                            format = "wmf";
                            break;
                        case Workbook.PICTURE_TYPE_PICT:
                            format = "pict";
                            break;*/
                        case Workbook.PICTURE_TYPE_JPEG:
                            format = "jpeg";
                            break;
                        case Workbook.PICTURE_TYPE_PNG:
                            format = "png";
                            break;
                        /*case Workbook.PICTURE_TYPE_DIB:
                            format = "dib";
                            break;*/
                        default:
                            throw new IllegalArgumentException("Picture type[" + type + "] is unsupported");
                    }
                }
                ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                ImageIO.write(((BufferedImagePicture) value).getBufferedImage(), format, byteArrayOut);
                createPicture(new ByteArrayPicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                        type, value.getFormat(), byteArrayOut.toByteArray()), anchor, workbook, drawing);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else if (value instanceof FilePicture) {
            try {
                String format;
                File file = ((FilePicture) value).getFile();
                Path path = Paths.get(file.getAbsolutePath());
                String mime = Files.probeContentType(path);
                switch (mime) {
                    /*case "PICTURE_TYPE_EMF":
                        type = Workbook.PICTURE_TYPE_EMF;
                        break;
                    case "PICTURE_TYPE_WMF":
                        type = Workbook.PICTURE_TYPE_WMF;
                        break;
                    case "PICTURE_TYPE_PICT":
                        type = Workbook.PICTURE_TYPE_PICT;
                        break;*/
                    case "image/jpeg":
                    case "image/pjpeg":
                        type = Workbook.PICTURE_TYPE_JPEG;
                        format = "jpeg";
                        break;
                    case "image/png":
                        type = Workbook.PICTURE_TYPE_PNG;
                        format = "png";
                        break;
                    /*case "PICTURE_TYPE_DIB":
                        type = Workbook.PICTURE_TYPE_DIB;
                        break;*/
                    default:
                        throw new IllegalArgumentException("Mime type[" + type + "] is unsupported");
                }
                /*BufferedImage bufferedImage = ImageIO.read(file);
                createPicture(new BufferedImagePicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                        type, value.getFormat(), bufferedImage), anchor, workbook, drawing);*/
                FileInputStream fis = new FileInputStream(file);
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                byte[] b = new byte[1024];
                int n;
                while ((n = fis.read(b)) != -1) {
                    bos.write(b, 0, n);
                }
                fis.close();
                bos.close();
                createPicture(new ByteArrayPicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                        type, format, bos.toByteArray()), anchor, workbook, drawing);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else if (value instanceof Base64Picture) {
            String format = value.getFormat();
            String jpeg = "data:image/jpeg;base64,";
            String png = "data:image/png;base64,";
            String base64 = ((Base64Picture) value).getBase64();
            if (base64.startsWith(jpeg)) {
                base64 = base64.replace(jpeg, "");
                type = Workbook.PICTURE_TYPE_JPEG;
                format = "jpeg";
            } else if (base64.startsWith(png)) {
                base64 = base64.replace(png, "");
                type = Workbook.PICTURE_TYPE_PNG;
                format = "png";
            }
            byte[] bytes = Base64.getDecoder().decode(base64);
            createPicture(new ByteArrayPicture(value.getPadding(), value.getLocation(), value.getAnchorType(),
                    type, format, bytes), anchor, workbook, drawing);
        }
    }
}
