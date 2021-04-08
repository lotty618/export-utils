package com.quanxi.qxexportutils.util.doc.spiredoc;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.BuiltinStyle;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.TextRange;

import java.awt.*;

public class TableofContents {
    public static void main(String[] args){
        //创建Word文档
        Document doc = new Document();
        //添加一个section
        Section section = doc.addSection();

        //添加段落
        Paragraph para = section.addParagraph();
        TextRange tr = para.appendText("Table of Contents");
        //设置字体大小和颜色
        tr.getCharacterFormat().setFontSize(11);
        tr.getCharacterFormat().setTextColor(Color.blue);
        //设置段后间距
        para.getFormat().setAfterSpacing(10);

        //添加段落
        para = section.addParagraph();
        //通过指定最低的Heading级别1和最高的Heading级别3，创建包含Heading 1、2、3，制表符前导符和右对齐页码的默认样式的Word目录。标题级别范围必须介于1到9之间
        para.appendTOC(1, 3);

        //添加一个section
        section = doc.addSection();
        //添加一个段落
        para = section.addParagraph();
        para.appendText("Heading 1");
        //应用Heading 1样式到段落
        para.applyStyle(BuiltinStyle.Heading_1);
        section.addParagraph();

        //添加一个段落
        para = section.addParagraph();
        para.appendText("Heading 2");
        //应用Heading 2样式到段落
        para.applyStyle(BuiltinStyle.Heading_2);
        section.addParagraph();

        //添加一个段落
        para = section.addParagraph();
        para.appendText("Heading 3");
        //应用Heading 3样式到段落
        para.applyStyle(BuiltinStyle.Heading_3);
        section.addParagraph();

        //另加
        //添加一个section
        section = doc.addSection();
        //添加一个段落
        para = section.addParagraph();
        para.appendText("Heading ** 1");
        //应用Heading 1样式到段落
        para.applyStyle(BuiltinStyle.Heading_1);
        section.addParagraph();

        for (int i = 0; i < 50; i++) {
            //添加一个section
            section = doc.addSection();
            //添加一个段落
            para = section.addParagraph();
            para.appendText("Heading ** " + i);
            //应用Heading 1样式到段落
            para.applyStyle(BuiltinStyle.Heading_1);
            section.addParagraph();
        }

        //更新目录
        doc.updateTableOfContents();

        //保存结果文档
        doc.saveToFile("createTableOfContents.docx", FileFormat.Docx);
    }
}
