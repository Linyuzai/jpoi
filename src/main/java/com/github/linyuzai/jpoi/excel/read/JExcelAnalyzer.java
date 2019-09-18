package com.github.linyuzai.jpoi.excel.read;

import com.github.linyuzai.jpoi.excel.JExcelBase;
import com.github.linyuzai.jpoi.excel.read.adapter.ReadAdapter;
import com.github.linyuzai.jpoi.excel.read.getter.ValueGetter;
import com.github.linyuzai.jpoi.excel.write.converter.ValueConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class JExcelAnalyzer extends JExcelBase<JExcelAnalyzer> {

    private Workbook workbook;
    private List<ValueConverter> valueConverters;
    private ValueGetter valueGetter;
    private ReadAdapter readAdapter;

    public JExcelAnalyzer(Workbook workbook) {
        this.workbook = workbook;
    }

    public Object analyze() {
        return analyze(workbook, readAdapter, valueConverters, valueGetter);
    }

    private Object analyze(Workbook workbook, ReadAdapter readAdapter, List<ValueConverter> valueConverters, ValueGetter valueGetter) {
        List<Object> sheetValues = new ArrayList<>();
        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            List<Object> rowValues = new ArrayList<>();
            for (int r = 0; r < sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                List<Object> cellValues = new ArrayList<>();
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c);
                    Object o = valueGetter.getValue(s, r, c, cell, row, sheet, workbook);
                    Object value = readAdapter.readCell(o, s, r, c, cell, row, sheet, workbook);
                    ValueConverter valueConverter = null;
                    for (ValueConverter vc : valueConverters) {
                        if (vc.supportValue(s, r, c, o)) {
                            valueConverter = vc;
                            break;
                        }
                    }
                    if (valueConverter == null) {
                        throw new RuntimeException("No value converter matched");
                    }
                    Object cellValue = valueConverter.convertValue(s, r, c, value);
                    cellValues.add(cellValue);
                }
                Object rowValue = readAdapter.readRow(cellValues, s, r, row, sheet, workbook);
                rowValues.add(rowValue);
            }
            Object sheetValue = readAdapter.readSheet(rowValues, s, sheet, workbook);
            sheetValues.add(sheetValue);
        }
        Object workbookValue = readAdapter.readWorkbook(sheetValues, workbook);

        return workbookValue;
    }
}
