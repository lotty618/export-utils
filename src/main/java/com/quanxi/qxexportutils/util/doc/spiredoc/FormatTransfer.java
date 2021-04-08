package com.quanxi.qxexportutils.util.doc.spiredoc;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;

public class FormatTransfer {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("test1.docx");
        document.saveToFile("TARGET test1.pdf", FileFormat.PDF);
    }
}
