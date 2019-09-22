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
        return 0;
    }

    @Override
    public int getRowIndex() {
        return 0;
    }

    @Override
    public Sheet getSheet() {
        return null;
    }

    @Override
    public Row getRow() {
        return null;
    }

    @Override
    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    @Override
    public void setBlank() {

    }

    @Override
    public CellType getCellType() {
        return cellType;
    }

    @Override
    public CellType getCellTypeEnum() {
        return null;
    }

    @Override
    public CellType getCachedFormulaResultType() {
        return null;
    }

    @Override
    public CellType getCachedFormulaResultTypeEnum() {
        return null;
    }

    @Override
    public void setCellValue(double value) {

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

    }

    @Override
    public CellStyle getCellStyle() {
        return null;
    }

    @Override
    public void setAsActiveCell() {

    }

    @Override
    public CellAddress getAddress() {
        return null;
    }

    @Override
    public void setCellComment(Comment comment) {

    }

    @Override
    public Comment getCellComment() {
        return null;
    }

    @Override
    public void removeCellComment() {

    }

    @Override
    public Hyperlink getHyperlink() {
        return null;
    }

    @Override
    public void setHyperlink(Hyperlink link) {

    }

    @Override
    public void removeHyperlink() {

    }

    @Override
    public CellRangeAddress getArrayFormulaRange() {
        return null;
    }

    @Override
    public boolean isPartOfArrayFormulaGroup() {
        return false;
    }
}
