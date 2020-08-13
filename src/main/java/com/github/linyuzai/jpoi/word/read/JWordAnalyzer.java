package com.github.linyuzai.jpoi.word.read;

import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

@Deprecated
public class JWordAnalyzer {

    private Document document;

    public JWordAnalyzer(Document document) {
        this.document = document;
    }

    public Document getDocument() {
        return document;
    }

    public void read() {
        if (document instanceof XWPFDocument) {
            XWPFDocument d = (XWPFDocument) document;
        }
    }
}
