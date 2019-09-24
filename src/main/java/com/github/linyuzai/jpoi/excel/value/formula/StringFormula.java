package com.github.linyuzai.jpoi.excel.value.formula;

public class StringFormula implements SupportFormula {

    private String formula;

    public StringFormula(String formula) {
        this.formula = formula;
    }

    @Override
    public String getFormula() {
        return formula;
    }

    @Override
    public void setFormula(String formula) {
        this.formula = formula;
    }
}
