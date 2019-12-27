package com.github.linyuzai.jpoi.excel.write;

import com.github.linyuzai.jpoi.exception.JPoiException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

public class JExcelWriter {
    private JExcelTransfer.Values values;

    public JExcelWriter(JExcelTransfer.Values values) {
        this.values = values;
    }

    public void to(OutputStream outputStream) throws IOException {
        to(outputStream, true);
    }

    public void to(OutputStream outputStream, boolean close) throws IOException {
        Workbook workbook = values.getWorkbook();
        workbook.write(outputStream);
        if (close) {
            workbook.close();
        }
    }

    public void to(File file) throws IOException {
        Workbook workbook = values.getWorkbook();
        String path = file.getAbsolutePath();
        if (workbook instanceof HSSFWorkbook && !path.endsWith(".xls")) {
            file = new File(path + ".xls");
        } else if (workbook instanceof XSSFWorkbook && !path.endsWith(".xlsx")) {
            file = new File(path + ".xlsx");
        }
        File p = file.getParentFile();
        if (!p.exists()) {
            if (!p.mkdirs()) {
                throw new JPoiException(p.getAbsolutePath() + " mkdirs failure");
            }
        }
        if (!file.exists()) {
            if (!file.createNewFile()) {
                throw new JPoiException(file.getAbsolutePath() + " create failure");
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        to(fileOutputStream);
        fileOutputStream.close();
    }

    public List<Throwable> getThrowableRecords() {
        return values.getThrowableRecords();
    }
}
