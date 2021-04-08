package com.quanxi.qxexportutils.util.doc.spiredoc;

import com.spire.doc.*;
import com.spire.doc.collections.ParagraphItemCollection;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.ParagraphBase;
import com.spire.doc.fields.TableOfContent;
import com.spire.doc.fields.TextRange;
import com.spire.doc.formatting.CharacterFormat;

import java.awt.*;

public class WordUtils {
    private Document document;
    private String fileName;
    private FileFormat format = FileFormat.Docx;

    public static final String STYLE_TITLE = "titleStyle";
    public static final String STYLE_TITLE_HEADER = "titleHeaderStyle";
    public static final String STYLE_PARA = "paraStyle";
    public static final String STYLE_PARA_BELOW = "textBelowStyle";
    public static final String STYLE_SUM_TABLE_TITLE = "styleSumTableTitle";
    public static final String STYLE_SUM_TABLE_CONTENT = "styleSumTableContent";
    public static final String STYLE_SUB_TABLE_TITLE = "styleSubTableTitle";
    public static final String STYLE_SUB_TABLE_CONTENT = "styleSubTableContent";

    public WordUtils(String saveFileName) {
        //创建Word文档
        document = new Document();
        this.fileName = saveFileName;

        initParaStyle();
    }

    /**
     * 合并文档到新文档
     * @param docFileName
     * @param saveFileName
     */
    public WordUtils(String docFileName, String saveFileName) {
        //创建Word文档
        document = new Document(docFileName);
        this.fileName = saveFileName;

        initParaStyle();
    }

    public Document getDocument() {
        return this.document;
    }

    /**
     * 初始化段落文本格式
     */
    private void initParaStyle() {
        //设置标题格式
        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("titleStyle");
        style1.getCharacterFormat().setBold(true);
//        style1.getCharacterFormat().setTextColor(Color.BLUE);
        style1.getCharacterFormat().setTextColor(new Color(54, 95, 145));
        style1.getCharacterFormat().setFontName("黑体");
        style1.getCharacterFormat().setFontSize(24f);
        style1.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        document.getStyles().add(style1);

        //设置标题格式
        ParagraphStyle style11 = new ParagraphStyle(document);
        style11.setName("titleHeaderStyle");
        style11.getCharacterFormat().setBold(true);
        style11.getCharacterFormat().setTextColor(Color.BLUE);
        style11.getCharacterFormat().setFontName("黑体");
        style11.getCharacterFormat().setFontSize(20f);
        style11.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        document.getStyles().add(style11);

        //设置段落的格式
        ParagraphStyle style2 = new ParagraphStyle(document);
        style2.setName("paraStyle");
        style2.getCharacterFormat().setFontName("仿宋");
        style2.getCharacterFormat().setFontSize(14f);
        style2.getParagraphFormat().setFirstLineIndent(20f);
        document.getStyles().add(style2);

        //设置表格下方文字格式
        ParagraphStyle style3 = new ParagraphStyle(document);
        style3.setName("textBelowStyle");
        style3.getCharacterFormat().setFontName("仿宋");
        style3.getCharacterFormat().setFontSize(14f);
        style3.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
        style3.getParagraphFormat().setFirstLineIndent(50f);
        document.getStyles().add(style3);

        //总表表格标题
        ParagraphStyle styleSumTableTitle = new ParagraphStyle(document);
        styleSumTableTitle.setName("styleSumTableTitle");
        styleSumTableTitle.getCharacterFormat().setFontName("仿宋");
        styleSumTableTitle.getCharacterFormat().setFontSize(12f);
        styleSumTableTitle.getCharacterFormat().setBold(true);
        styleSumTableTitle.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        document.getStyles().add(styleSumTableTitle);

        //总表表格内容
        ParagraphStyle styleSumTableContent = new ParagraphStyle(document);
        styleSumTableContent.setName("styleSumTableContent");
        styleSumTableContent.getCharacterFormat().setFontName("仿宋");
        styleSumTableContent.getCharacterFormat().setFontSize(14f);
        styleSumTableContent.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        document.getStyles().add(styleSumTableContent);

        //分表表格标题
        ParagraphStyle styleSubTableTitle = new ParagraphStyle(document);
        styleSubTableTitle.setName("styleSubTableTitle");
        styleSubTableTitle.getCharacterFormat().setFontName("黑体");
        styleSubTableTitle.getCharacterFormat().setFontSize(14f);
        styleSubTableTitle.getCharacterFormat().setBold(true);
        styleSubTableTitle.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        document.getStyles().add(styleSubTableTitle);

        //分表表格内容
        ParagraphStyle styleSubTableContent = new ParagraphStyle(document);
        styleSubTableContent.setName("styleSubTableContent");
        styleSubTableContent.getCharacterFormat().setFontName("仿宋");
        styleSubTableContent.getCharacterFormat().setFontSize(12f);
        document.getStyles().add(styleSubTableContent);
    }

    /**
     * 创建Section
     * @return
     */
    public Section createSection() {
        //添加一个section
        return document.addSection();
    }

    /**
     * 插入段落
     * @param section
     * @param text
     * @param style
     * @param isNewPage
     * @return
     */
    public Paragraph insertParagraph(Section section, String text, String style) {
        //添加段落至section
        Paragraph para = section.addParagraph();
        para.appendText(text);
        para.applyStyle(style);

        return para;
    }

    public Paragraph insertParagraph(Section section, String text, BuiltinStyle builtinStyle) {
        //添加段落至section
        Paragraph para = section.addParagraph();
        para.appendText(text);
        para.applyStyle(builtinStyle);

        return para;
    }

    /**
     * 使用BuiltinStyle builtinStyle配合TOC域{/o "1-3" /h /z /u}自动生成目录
     * 如果使用自定义样式或标签，则无法自动更新域，文档上显示TOC，需要右键更新域才能看到目录，所以此处使用默认标题样式BuiltinStyle.Heading_1，
     * 这样配合TOC域{/o "1-3" /h /z /u}才会自动更新域，然后在此样式基础上做定制化的修改，如设置字体、大小、颜色、行间距等。
     * @param section
     * @param text
     * @param builtinStyle
     * @param style
     * @return
     */
    public Paragraph insertParagraph(Section section, String text, BuiltinStyle builtinStyle, String style) {
        //添加段落至section
        Paragraph para = insertParagraph(section, text, builtinStyle);

        CharacterFormat cf = para.getItems().get(0).getCharacterFormat();
        cf.setFontName("黑体");
        cf.setFontSize(24f);
        cf.setBold(true);
        cf.setTextColor(Color.RED);
        para.getFormat().setLineSpacing(2f);
        para.getItems().get(0).applyCharacterFormat(cf);

        return para;
    }

    /**
     * 插入图片
     * @param section
     * @param filePath
     * @param width
     * @param height
     * @param isNewPage
     * @return
     */
    public DocPicture insertPicture(Section section, String filePath, float width, float height, boolean isNewPage) {
        //添加段落
        Paragraph para = section.addParagraph();

        //添加图片到段落
        DocPicture picture = para.appendPicture(filePath);
        //设置图片宽度
        picture.setWidth(width);
        //设置图片高度
        picture.setHeight(height);

//        if (isNewPage) {
//            para.insertSectionBreak(SectionBreakType.New_Page);
//        }

        return picture;
    }

    /**
     * 插入表格
     * @param section
     * @param header
     * @param data
     * @return
     */
    public Table insertTable(Section section, String[] header, String[][] data, String headerStyle, String contentStyle) {
        //添加表格
        Table table = section.addTable(true);
        //设置表格的行数和列数
        table.resetCells(data.length + 1, header.length);

        //设置第一行作为表格的表头并添加数据
        TableRow row = table.getRows().get(0);
        row.isHeader(true);
        row.setHeight(20);
        row.setHeightType(TableRowHeightType.Exactly);
//        row.getRowFormat().setBackColor(Color.gray);
        for (int i = 0; i < header.length; i++) {
            row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            Paragraph p = row.getCells().get(i).addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange range1 = p.appendText(header[i]);

            range1.getCharacterFormat().setFontName("仿宋");
            range1.getCharacterFormat().setFontSize(12f);
            range1.getCharacterFormat().setBold(true);
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//            range1.applyStyle(headerStyle);
        }

        //添加数据到剩余行
        for (int r = 0; r < data.length; r++) {
            TableRow dataRow = table.getRows().get(r + 1);
//            dataRow.setHeight(25);
//            dataRow.setHeightType(TableRowHeightType.Exactly);
            dataRow.getRowFormat().setBackColor(Color.white);
            for (int c = 0; c < data[r].length; c++) {
                TableCell tableCell = dataRow.getCells().get(c);
                tableCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                Paragraph pCell = tableCell.addParagraph();
                TextRange range2 = pCell.appendText(data[r][c]);

                range2.getCharacterFormat().setFontName("仿宋");
                range2.getCharacterFormat().setFontSize(14f);
                pCell.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//                range2.applyStyle(contentStyle);
            }
        }

        return table;
    }


    /**
     * 插入表格
     * @param section
     * @param header
     * @param width 每列宽度
     * @param height 标题行高度
     * @param data
     * @return
     */
    public Table insertTable(Section section, float[] widths, float height, String[] header, String[][] data, String headerStyle, String contentStyle) {
        //添加表格
        Table table = section.addTable(true);
        //设置表格的行数和列数
        table.resetCells(data.length + 1, header.length);

        //设置第一行作为表格的表头并添加数据
        TableRow row = table.getRows().get(0);
        row.isHeader(true);
        row.setHeight(20);
        row.setHeightType(TableRowHeightType.Exactly);
//        row.getRowFormat().setBackColor(Color.gray);
        for (int i = 0; i < header.length; i++) {
            row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            Paragraph p = row.getCells().get(i).addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange range1 = p.appendText(header[i]);

            range1.getCharacterFormat().setFontName("仿宋");
            range1.getCharacterFormat().setFontSize(12f);
            range1.getCharacterFormat().setBold(true);
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//            range1.applyStyle(headerStyle);
        }

        //添加数据到剩余行
        for (int r = 0; r < data.length; r++) {
            TableRow dataRow = table.getRows().get(r + 1);
//            dataRow.setHeight(25);
//            dataRow.setHeightType(TableRowHeightType.Exactly);
            dataRow.getRowFormat().setBackColor(Color.white);
            for (int c = 0; c < data[r].length; c++) {
                TableCell tableCell = dataRow.getCells().get(c);
                tableCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                Paragraph pCell = tableCell.addParagraph();
                TextRange range2 = pCell.appendText(data[r][c]);

                range2.getCharacterFormat().setFontName("仿宋");
                range2.getCharacterFormat().setFontSize(14f);
                pCell.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//                range2.applyStyle(contentStyle);
            }
        }

        return table;
    }

    /**
     * 替换文档中的文本
     * @param src
     * @param dest
     */
    public void replaceText(String src, String dest) {
        document.replace(src, dest, false, true);
    }

    /**
     * 添加目录
     * @param para
     */
    public void addTableOfContents(Paragraph para) {
        //使用TOC域开关，创建一个包含Heading 1、2、3但省略了页码的自定义目录
//        TableOfContent toc = new TableOfContent(document, "{\\o \"1-3\" \\h \\z \\w}");       //"{\\t \"titleHeaderStyle,1\"}"
//        TableOfContent toc = new TableOfContent(document, "{\\b \"TOC_1\"}");
        para.appendTOC(1,3);  //默认TOC，同 /o "1-3" /h /z /u

        //设置目录样式
//        ParagraphStyle tocStyle1 = (ParagraphStyle) Style.createBuiltinStyle(BuiltinStyle.Toc_1, document);
//        tocStyle1.getCharacterFormat().setFontSize(10.5f);
//        tocStyle1.getParagraphFormat().setLineSpacing(17f);
//        document.getStyles().add(tocStyle1);

//        para.getItems().add(toc);
//        para.appendFieldMark(FieldMarkType.Field_Separator);
//        para.appendText("TOC");
//        para.appendFieldMark(FieldMarkType.Field_End);
//        document.setTOC(toc);
    }

    /**
     * 生成目录
     */
    public void updateTableOfContents() {
//        document.isUpdateFields(true);
        document.updateTableOfContents();
    }

    /**
     * 写入文件
     */
    public void save() {
        document.saveToFile(fileName, format);
    }
}
