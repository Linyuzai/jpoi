package com.github.linyuzai.jpoi.excel.value.formula;

public class FormulaValue implements SupportFormula {

    private String formula;

    private Object value;

    public FormulaValue(String formula, Object value) {
        this.formula = formula;
        this.value = value;
    }

    @Override
    public String getFormula() {
        return formula;
    }

    @Override
    public Object getValue() {
        return value;
    }

    //@Override
    public void setFormula(String formula) {
        this.formula = formula;
    }

    public void setValue(Object value) {
        this.value = value;
    }
}
