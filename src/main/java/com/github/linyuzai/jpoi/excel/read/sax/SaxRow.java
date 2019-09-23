package com.github.linyuzai.jpoi.excel.read.sax;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class SaxRow implements Row {

    private int rowNum;
    private List<Cell> cells = new ArrayList<>();

    public List<Cell> getCells() {
        return cells;
    }

    public void setCells(List<Cell> cells) {
        this.cells = cells;
    }

    @Override
    public Cell createCell(int column) {
        return new SaxCell();
    }

    @Override
    public Cell createCell(int column, CellType type) {
        Cell cell = createCell(column);
        cell.setCellType(type);
        return cell;
    }

    @Override
    public void removeCell(Cell cell) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    @Override
    public int getRowNum() {
        return rowNum;
    }

    @Override
    public Cell getCell(int cellNum) {
        return cells.get(cellNum);
    }

    @Override
    public Cell getCell(int cellNum, MissingCellPolicy policy) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getFirstCellNum() {
        return 0;
    }

    @Override
    public short getLastCellNum() {
        return (short) (cells.size() - 1);
    }

    @Override
    public int getPhysicalNumberOfCells() {
        return cells.size();
    }

    @Override
    public void setHeight(short height) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setZeroHeight(boolean zHeight) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getZeroHeight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHeightInPoints(float height) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getHeight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public float getHeightInPoints() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isFormatted() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellStyle getRowStyle() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowStyle(CellStyle style) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Cell> cellIterator() {
        return cells.iterator();
    }

    @Override
    public Sheet getSheet() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getOutlineLevel() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Cell> iterator() {
        return cells.iterator();
    }
}
