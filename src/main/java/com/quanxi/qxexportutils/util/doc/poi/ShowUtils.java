package com.quanxi.qxexportutils.util.doc.poi;

import com.quanxi.qxexportutils.util.doc.poi.word2html.ExportHtml;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class ShowUtils {
    public static void main(String[] args) {
        try {
            String ret = ExportHtml.doc2Html(System.getProperty("user.dir"), "testpoi20", ".docx");
            System.out.println(ret);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
