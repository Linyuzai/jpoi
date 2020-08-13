package com.github.linyuzai.jpoi.excel.handler;

public class ExceptionValue {

    private int sheet;
    private int row;
    private int cell;
    private Throwable throwable;

    public ExceptionValue(int sheet, int row, int cell, Throwable throwable) {
        this.sheet = sheet;
        this.row = row;
        this.cell = cell;
        this.throwable = throwable;
    }

    public int getSheet() {
        return sheet;
    }

    public void setSheet(int sheet) {
        this.sheet = sheet;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCell() {
        return cell;
    }

    public void setCell(int cell) {
        this.cell = cell;
    }

    public Throwable getThrowable() {
        return throwable;
    }

    public void setThrowable(Throwable throwable) {
        this.throwable = throwable;
    }
}
