package com.github.linyuzai.jpoi.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class JExcelWriter {
    private Workbook workbook;

    public JExcelWriter(Workbook workbook) {
        if (workbook == null) {
            throw new RuntimeException("Workbook is null");
        }
        this.workbook = workbook;
    }

    public void to(OutputStream outputStream) throws IOException {
        to(outputStream, true);
    }

    public void to(OutputStream outputStream, boolean close) throws IOException {
        workbook.write(outputStream);
        if (close) {
            workbook.close();
        }
    }

    public void to(File file) throws IOException {
        String path = file.getAbsolutePath();
        if (workbook instanceof HSSFWorkbook && !path.endsWith(".xls")) {
            file = new File(path + ".xls");
        } else if (workbook instanceof XSSFWorkbook && !path.endsWith(".xlsx")) {
            file = new File(path + ".xlsx");
        }
        File p = file.getParentFile();
        if (!p.exists()) {
            if (!p.mkdirs()) {
                throw new RuntimeException(p.getAbsolutePath() + " mkdirs failure");
            }
        }
        if (!file.exists()) {
            if (!file.createNewFile()) {
                throw new RuntimeException(file.getAbsolutePath() + " create failure");
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        to(fileOutputStream);
        fileOutputStream.close();
    }
}
