package com.github.linyuzai.jpoi.excel.value.formula;

import com.github.linyuzai.jpoi.support.SupportValue;

public interface SupportFormula extends SupportValue {

    String getFormula();

    Object getValue();
    //void setFormula(String formula);
}
