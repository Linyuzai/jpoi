package com.github.linyuzai.jpoi.excel.read.sax;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

public class SaxCreationHelper implements CreationHelper {

    private SaxFormulaEvaluator formulaEvaluator = new SaxFormulaEvaluator();

    @Override
    public RichTextString createRichTextString(String text) {
        return null;
    }

    @Override
    public DataFormat createDataFormat() {
        return null;
    }

    @Override
    public Hyperlink createHyperlink(HyperlinkType type) {
        return null;
    }

    @Override
    public FormulaEvaluator createFormulaEvaluator() {
        return formulaEvaluator;
    }

    @Override
    public ExtendedColor createExtendedColor() {
        return null;
    }

    @Override
    public ClientAnchor createClientAnchor() {
        return null;
    }

    @Override
    public AreaReference createAreaReference(String reference) {
        return null;
    }

    @Override
    public AreaReference createAreaReference(CellReference topLeft, CellReference bottomRight) {
        return null;
    }
}
