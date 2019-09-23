package com.github.linyuzai.jpoi.excel.read.sax;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Calendar;
import java.util.Date;

public class SaxCell implements Cell {

    private Object cellValue;
    private CellType cellType;

    @Override
    public int getColumnIndex() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getRowIndex() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Sheet getSheet() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Row getRow() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    @Override
    public void setBlank() {
        cellValue = "";
    }

    @Override
    public CellType getCellType() {
        return cellType;
    }

    @Override
    public CellType getCellTypeEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellType getCachedFormulaResultType() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellType getCachedFormulaResultTypeEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setCellValue(double value) {
        this.cellValue = value;
    }

    @Override
    public void setCellValue(Date value) {
        this.cellValue = value;
    }

    @Override
    public void setCellValue(Calendar value) {
        this.cellValue = value;
    }

    @Override
    public void setCellValue(RichTextString value) {
        this.cellValue = value;
    }

    @Override
    public void setCellValue(String value) {
        this.cellValue = value;
    }

    @Override
    public void setCellFormula(String formula) throws FormulaParseException, IllegalStateException {
        this.cellValue = formula;
    }

    @Override
    public void removeFormula() throws IllegalStateException {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getCellFormula() {
        return (String) cellValue;
    }

    @Override
    public double getNumericCellValue() {
        return (double) cellValue;
    }

    @Override
    public Date getDateCellValue() {
        return (Date) cellValue;
    }

    @Override
    public RichTextString getRichStringCellValue() {
        return (RichTextString) cellValue;
    }

    @Override
    public String getStringCellValue() {
        return (String) cellValue;
    }

    @Override
    public void setCellValue(boolean value) {
        this.cellValue = value;
    }

    @Override
    public void setCellErrorValue(byte value) {
        this.cellValue = value;
    }

    @Override
    public boolean getBooleanCellValue() {
        return (boolean) cellValue;
    }

    @Override
    public byte getErrorCellValue() {
        return (byte) cellValue;
    }

    @Override
    public void setCellStyle(CellStyle style) {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellStyle getCellStyle() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setAsActiveCell() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellAddress getAddress() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setCellComment(Comment comment) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Comment getCellComment() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeCellComment() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Hyperlink getHyperlink() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHyperlink(Hyperlink link) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeHyperlink() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellRangeAddress getArrayFormulaRange() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isPartOfArrayFormulaGroup() {
        throw new UnsupportedOperationException();
    }
}
