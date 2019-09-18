package com.github.linyuzai.jpoi.excel.write.value;

import com.github.linyuzai.jpoi.support.SupportValue;

public interface SupportFormula extends SupportValue {

    String getFormula();

    void setFormula(String formula);
}
