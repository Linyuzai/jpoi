package com.github.linyuzai.jpoi.excel.holder;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import org.apache.poi.ss.usermodel.*;

public class ValueHolder<S, V> {

    private int sheetIndex;
    private int rowIndex;
    private int cellIndex;
    private Sheet sheet;
    private Row row;
    private Cell cell;
    private Drawing<?> drawing;
    private CreationHelper creationHelper;
    private S source;
    private V value;
    private ValueConverter valueConverter;

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    public Drawing<?> getDrawing() {
        return drawing;
    }

    public void setDrawing(Drawing<?> drawing) {
        this.drawing = drawing;
    }

    public CreationHelper getCreationHelper() {
        return creationHelper;
    }

    public void setCreationHelper(CreationHelper creationHelper) {
        this.creationHelper = creationHelper;
    }

    public S getSource() {
        return source;
    }

    public void setSource(S source) {
        this.source = source;
    }

    public V getValue() {
        return value;
    }

    public void setValue(V value) {
        this.value = value;
    }

    public ValueConverter getValueConverter() {
        return valueConverter;
    }

    public void setValueConverter(ValueConverter valueConverter) {
        this.valueConverter = valueConverter;
    }
}
