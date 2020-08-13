package com.github.linyuzai.jpoi.word;

import com.github.linyuzai.jpoi.exception.JPoiException;
import com.github.linyuzai.jpoi.word.read.JWordAnalyzer;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;

@Deprecated
public class JWord {

    public static JWordAnalyzer docx(InputStream is) throws IOException {
        return from(new XWPFDocument(is));
    }

    public static JWordAnalyzer from(Document document) {
        if (document == null) {
            throw new JPoiException("Document is null");
        }
        return new JWordAnalyzer(document);
    }
}
