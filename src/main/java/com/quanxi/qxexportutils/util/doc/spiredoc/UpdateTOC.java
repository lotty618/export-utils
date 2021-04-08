package com.quanxi.qxexportutils.util.doc.spiredoc;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.BreakType;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.SectionBreakType;
import com.spire.doc.fields.TableOfContent;
import com.spire.doc.fields.TextRange;

import java.awt.*;

public class UpdateTOC {
    public static void main(String[] args) {
        Document doc = new Document("test1.docx");

        //在文档最前面插入一个段落，写入文本并格式化
        Paragraph parainserted = new Paragraph(doc);
        TextRange tr= parainserted.appendText("目 录");
        tr.getCharacterFormat().setBold(true);
        tr.getCharacterFormat().setTextColor(Color.gray);
        System.out.println(doc.getSections().getCount());
        doc.getSections().get(1).getParagraphs().insert(0,parainserted);
        parainserted.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //通过域代码添加目录表
        TableOfContent toc = new TableOfContent(doc, "{\\o \"1-3\" \\h \\z \\u}");
        Section section = doc.getSections().get(1);
        section.addColumn(100f, 20f);
        section.addColumn(100f, 20f);

        section.getParagraphs().get(0).appendTOC(1,3);
        section.getParagraphs().get(0).appendBreak(BreakType.Page_Break);
        doc.updateTableOfContents();

        //保存文档
        doc.saveToFile("test2.docx", FileFormat.Docx_2010);
    }
}
