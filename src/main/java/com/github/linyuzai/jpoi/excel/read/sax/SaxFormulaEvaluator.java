package com.github.linyuzai.jpoi.excel.read.sax;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.util.Map;

public class SaxFormulaEvaluator implements FormulaEvaluator {
    @Override
    public void clearAllCachedResultValues() {

    }

    @Override
    public void notifySetFormula(Cell cell) {

    }

    @Override
    public void notifyDeleteCell(Cell cell) {

    }

    @Override
    public void notifyUpdateCell(Cell cell) {

    }

    @Override
    public void evaluateAll() {

    }

    @Override
    public CellValue evaluate(Cell cell) {
        return new CellValue(cell.getStringCellValue());
    }

    @Override
    public CellType evaluateFormulaCell(Cell cell) {
        return null;
    }

    @Override
    public CellType evaluateFormulaCellEnum(Cell cell) {
        return null;
    }

    @Override
    public Cell evaluateInCell(Cell cell) {
        return null;
    }

    @Override
    public void setupReferencedWorkbooks(Map<String, FormulaEvaluator> workbooks) {

    }

    @Override
    public void setIgnoreMissingWorkbooks(boolean ignore) {

    }

    @Override
    public void setDebugEvaluationOutputForNextEval(boolean value) {

    }
}
