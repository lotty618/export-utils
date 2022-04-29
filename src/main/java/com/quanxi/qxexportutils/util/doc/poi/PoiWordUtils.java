package com.quanxi.qxexportutils.util.doc.poi;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.quanxi.qxexportutils.common.DocConst;
import com.quanxi.qxexportutils.util.doc.poi.export2word.CustomTOC;
import com.quanxi.qxexportutils.util.doc.poi.export2word.TextStyle;
import com.quanxi.qxexportutils.util.doc.poi.export2word.XWPFHelper;
import com.quanxi.qxexportutils.util.doc.poi.export2word.XWPFHelperTable;
import com.quanxi.qxexportutils.util.doc.poi.export2word.chart.ColumnChart;
import com.quanxi.qxexportutils.util.doc.poi.export2word.chart.LineChart;
import com.quanxi.qxexportutils.util.doc.poi.export2word.chart.PieChart;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;

public class PoiWordUtils {
    private XWPFHelperTable xwpfHelperTable = null;
    private XWPFHelper xwpfHelper = null;

    private XWPFDocument document;                      //文档对象
    private String fileName;                            //文件名
    private PAGE_ORIENTATION orientation;               //纸张方向

    public static TextStyle STYLE_TITLE;                //标题样式
    public static TextStyle STYLE_TOC;                  //目录样式
    public static TextStyle STYLE_DEFAULT;              //文本默认样式
    public static TextStyle STYLE_TABLE_SUM_HEADER;     //总表标题样式
    public static TextStyle STYLE_TABLE_SUM_CONTENT;    //总表内容样式
    public static TextStyle STYLE_TABLE_SUB_HEADER;     //分表标题样式
    public static TextStyle STYLE_TABLE_SUB_CONTENT;    //分表内容样式
    public static TextStyle STYLE_HEADER;               //页眉
    public static TextStyle STYLE_FOOTER;               //页脚

    public static final long PAGE_A4_WIDTH = 11906;     //A4纸宽度
    public static final long PAGE_A4_HEIGHT = 16838;    //A4纸高度

    public enum PAGE_ORIENTATION {
        HORIZONTAL,     //纸张：横向
        VERTICAL        //纸张：纵向
    }

    /**
     * 样式初始化
     */
    static {
        STYLE_TITLE = new TextStyle();
        STYLE_TITLE.setBold(true);
        STYLE_TITLE.setColorVal("000000");
        STYLE_TITLE.setFontSize("24");
        STYLE_TITLE.setFontFamily("黑体");

        STYLE_TOC = new TextStyle();
        STYLE_TOC.setColorVal("000000");
        STYLE_TOC.setFontSize("16");
        STYLE_TOC.setFontFamily("宋体");

        STYLE_DEFAULT = new TextStyle();
        STYLE_DEFAULT.setColorVal("000000");
        STYLE_DEFAULT.setFontSize("10");
        STYLE_DEFAULT.setFontFamily("等线");

        STYLE_TABLE_SUM_HEADER = new TextStyle();
        STYLE_TABLE_SUM_HEADER.setColorVal("000000");
        STYLE_TABLE_SUM_HEADER.setFontSize("12");
        STYLE_TABLE_SUM_HEADER.setFontFamily("仿宋");
        STYLE_TABLE_SUM_HEADER.setBold(true);

        STYLE_TABLE_SUM_CONTENT = new TextStyle();
        STYLE_TABLE_SUM_CONTENT.setColorVal("000000");
        STYLE_TABLE_SUM_CONTENT.setFontSize("14");
        STYLE_TABLE_SUM_CONTENT.setFontFamily("仿宋");

        STYLE_TABLE_SUB_HEADER = new TextStyle();
        STYLE_TABLE_SUB_HEADER.setColorVal("000000");
        STYLE_TABLE_SUB_HEADER.setFontSize("14");
        STYLE_TABLE_SUB_HEADER.setFontFamily("黑体");
        STYLE_TABLE_SUB_HEADER.setBold(true);

        STYLE_TABLE_SUB_CONTENT = new TextStyle();
        STYLE_TABLE_SUB_CONTENT.setColorVal("000000");
        STYLE_TABLE_SUB_CONTENT.setFontSize("12");
        STYLE_TABLE_SUB_CONTENT.setFontFamily("仿宋");

        STYLE_HEADER = new TextStyle();
        STYLE_HEADER.setColorVal("000000");
        STYLE_HEADER.setFontSize("10");
        STYLE_HEADER.setFontFamily("仿宋");
        STYLE_FOOTER = STYLE_HEADER;
    }

    /**
     * 构造函数
     * @param saveFileName  生成的文档路径
     */
    public PoiWordUtils(String saveFileName, PAGE_ORIENTATION orientation) {
        this.document = new XWPFDocument();
        this.fileName = saveFileName;
        this.orientation = orientation;
        this.xwpfHelperTable = new XWPFHelperTable();
        this.xwpfHelper = new XWPFHelper();

        //此处添加标题样式，对应TOC域的目录级别
        addCustomHeadingStyle("heading 1", 1);
    }

    /**
     * 构造函数
     * @param srcFileName   要加载的文档路径
     * @param saveFileName  生成的文档路径
     */
    public PoiWordUtils(String srcFileName, String saveFileName, PAGE_ORIENTATION orientation) {
        try {
            FileInputStream is = new FileInputStream(srcFileName);
            OPCPackage open = OPCPackage.open(is);

            this.document = new XWPFDocument(open);
            this.fileName = saveFileName;
            this.orientation = orientation;
            this.xwpfHelperTable = new XWPFHelperTable();
            this.xwpfHelper = new XWPFHelper();

            addCustomHeadingStyle("heading 1", 1);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 添加自定义标题样式
     * @param strStyleId    样式名称
     * @param headingLevel  标题级别
     */
    private void addCustomHeadingStyle(String strStyleId, int headingLevel ) {
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = document.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }

    public XWPFDocument getDocument() {
        return document;
    }

    /**
     * 添加段落
     * @param text          段落文本
     * @param isNewPage     是否在新页添加
     * @param style         文本样式
     * @param isTOCSet      是否设置为目录标题，True表示该段落将入目录
     * @return
     */
    public XWPFParagraph createParagrah(String text, boolean isNewPage, TextStyle style, boolean isTOCSet) {
        return createParagrah(orientation, text, isNewPage, style, isTOCSet);
    }


    /**
     * 添加段落
     * @param text          段落文本
     * @param isNewPage     是否在新页添加
     * @param style         文本样式
     * @param isTOCSet      是否设置为目录标题，True表示该段落将入目录
     * @return
     */
    public XWPFParagraph createParagrah(PAGE_ORIENTATION orientation, String text, boolean isNewPage, TextStyle style, boolean isTOCSet) {
        //设置纸张方向
        setPageOrientation(orientation);

        XWPFParagraph para = document.createParagraph();
        para.setPageBreak(isNewPage);   //是否创建新页
        para.setAlignment(ParagraphAlignment.CENTER);

        //添加文本
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setBold(style.isBold());
        run.setColor(style.getColorVal());
        run.setFontSize(Integer.parseInt(style.getFontSize()));

        if (isTOCSet) {
            //添加标题样式，对应为自定义的标题名称
            para.setStyle("heading 1");
        }

//        CTStyle ctStyle = CTStyle.Factory.newInstance();
//        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
//        onoffnull.setVal(STOnOff.ON);
//        ctStyle.getPPr().getRPr().setB(onoffnull);
//        XWPFStyle xwpfStyle = new XWPFStyle(ctStyle);
//        document.getStyles().addStyle(xwpfStyle);

        return para;
    }

    private void getStyleValue() {
        try {
            CTStyles styles = document.getStyle();
            List<CTStyle> styleArray = styles.getStyleList();
            for (CTStyle style : styleArray) {
                System.out.println(style.getName());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成表格
     * @param header    表格标题列表
     * @param data      表格数据
     * @param widths    每列宽度列表
     * @param height    每行高度
     * @return
     */
    public XWPFTable createTable(String[] header, TextStyle headerStyle, String[][] data, TextStyle contentStyle, long[] widths, long height) {
        XWPFTable table = document.createTable(data.length + 1, header.length);

        //设置表格样式
        List<XWPFTableRow> rowList = table.getRows();
        for(int i = 0; i < rowList.size(); i++) {
            XWPFTableRow infoTableRow = rowList.get(i);
            List<XWPFTableCell> cellList = infoTableRow.getTableCells();
            for(int j = 0; j < cellList.size(); j++) {
                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);

                XWPFRun cellParagraphRun = cellParagraph.createRun();
                TextStyle textStyle = STYLE_DEFAULT;

                xwpfHelperTable.setColumnWidthAndHAlign(cellList.get(j), widths[j]);

                if (i == 0) {
                    //标题
                    cellParagraphRun.setText(header[j]);
                    textStyle = headerStyle;
                } else {
                    //内容
                    cellParagraphRun.setText(data[i - 1][j]);
                    textStyle = contentStyle;
                }

                //设置文本样式
                cellParagraphRun.setBold(textStyle.isBold());
                cellParagraphRun.setColor(textStyle.getColorVal());
                cellParagraphRun.setFontSize(Integer.parseInt(textStyle.getFontSize()));
                cellParagraphRun.setFontFamily(textStyle.getFontFamily());
            }
        }
//        xwpfHelperTable.setTableHeight(table, 560, STVerticalJc.CENTER);
        xwpfHelperTable.setColumnHeight(table, height);

        return table;
    }

    /**
     * 生成表格
     * @param header    表格标题列表
     * @param data      表格数据
     * @param widths    每列宽度列表
     * @param height    每行高度
     * @param mergeCols 自动合并列索引项（从0开始，0表示第1列），如遇相同内容则合并单元格
     * @param mergeRows 自动合并行索引项（从0开始，0表示第1行），如遇相同内容则合并单元格
     * @return
     */
    public XWPFTable createTable(String[] header, TextStyle headerStyle, String[][] data, TextStyle contentStyle,
                                 long[] widths, long height, int[] mergeCols, int[] mergeRows) {
        XWPFTable table = document.createTable(data.length + 1, header.length);

        //设置表格样式
        List<XWPFTableRow> rowList = table.getRows();
        for(int i = 0; i < rowList.size(); i++) {
            XWPFTableRow infoTableRow = rowList.get(i);
            List<XWPFTableCell> cellList = infoTableRow.getTableCells();
            for(int j = 0; j < cellList.size(); j++) {
                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);

                XWPFRun cellParagraphRun = cellParagraph.createRun();
                TextStyle textStyle = STYLE_DEFAULT;

                xwpfHelperTable.setColumnWidthAndHAlign(cellList.get(j), widths[j]);

                if (i == 0) {
                    //标题
                    cellParagraphRun.setText(header[j]);
                    textStyle = headerStyle;
                } else {
                    //内容
                    cellParagraphRun.setText(data[i - 1][j]);
                    textStyle = contentStyle;
                }

                //判断是否该列属性可合并列
                mergeCellsVertically(table, mergeCols, i, j);
                //判断是否该行属性可合并列
                mergeCellsHorizontal(table, mergeRows, i, j);

                //设置文本样式
                cellParagraphRun.setBold(textStyle.isBold());
                cellParagraphRun.setColor(textStyle.getColorVal());
                cellParagraphRun.setFontSize(Integer.parseInt(textStyle.getFontSize()));
                cellParagraphRun.setFontFamily(textStyle.getFontFamily());
            }
        }
//        xwpfHelperTable.setTableHeight(table, 560, STVerticalJc.CENTER);
        xwpfHelperTable.setColumnHeight(table, height);

        return table;
    }

    /**
     * 跨列合并单元格
     * @param table     表格
     * @param mergeRows 自动合并行索引项（从0开始，0表示第1行），如遇相同内容则合并单元格
     * @param row       当前所在行
     * @param col       当前所在列
     */
    private void mergeCellsHorizontal(XWPFTable table, int[] mergeRows, int row, int col) {
        //获取当前单元格的文本
        String text = table.getRow(row).getCell(col).getText();
        int fromCol = col;

        //判断该行是否属于可合并行，并且保证当前列至少在第2行（加上标题）
        if (ArrayUtils.contains(mergeRows, row) && col > 0) {
            for (int j = col - 1; j >= 0; j--) {
                if (text.equals(table.getRow(row).getCell(j).getText())) {
                    fromCol = j;
                } else {
                    break;
                }
            }

            if (fromCol < col) {
                xwpfHelperTable.mergeCellsHorizontal(table, row, fromCol, col);
            }
        }
    }

    /**
     * 跨行合并单元格
     * @param table     表格
     * @param mergeCols 自动合并列索引项（从0开始，0表示第1列），如遇相同内容则合并单元格
     * @param row       当前所在行
     * @param col       当前所在列
     */
    private void mergeCellsVertically(XWPFTable table, int[] mergeCols, int row, int col) {
        //获取当前单元格的文本
        String text = table.getRow(row).getCell(col).getText();
        int fromRow = row;

        //判断该列是否属于可合并列，并且保证当前行至少在第3行（加上标题）
        if (ArrayUtils.contains(mergeCols, col) && row > 1) {
            for (int i = row - 1; i > 0; i--) {
                if (text.equals(table.getRow(i).getCell(col).getText())) {
                    fromRow = i;
                } else {
                    break;
                }
            }

            if (fromRow < row) {
                xwpfHelperTable.mergeCellsVertically(table, col, fromRow, row);
            }
        }
    }

    public void createControl(DocConst.CONTROL_TYPE type, Map<String, String> data) {
        XWPFParagraph para = document.createParagraph();
        XWPFRun run = para.createRun();

        run.setText("内容控件：[");

        CTSdtRun ctSdtRun = para.getCTP().addNewSdt();
        CTSdtPr sdtPr = ctSdtRun.addNewSdtPr();

        CTSdtListItem listItem;

        switch (type) {
            case DROPDOWNLIST:
                CTSdtDropDownList dropDownList = sdtPr.addNewDropDownList();
                listItem = dropDownList.addNewListItem();
                listItem.setDisplayText("请选择");
                listItem.setValue("请选择");

                boolean bIsSet = false;

                for (String s : data.keySet()) {
                    String val = data.get(s);
                    listItem = dropDownList.addNewListItem();
                    listItem.setDisplayText(val);
                    listItem.setValue(s);

                }
                break;
            case COMBOBOX:
                CTSdtComboBox comboBox = sdtPr.addNewComboBox();
                listItem = comboBox.addNewListItem();
                listItem.setDisplayText("请选择");
                listItem.setValue("请选择");

                for (String s : data.keySet()) {
                    String val = data.get(s);
                    listItem = comboBox.addNewListItem();
                    listItem.setDisplayText(val);
                    listItem.setValue(s);

                }

                break;
            case DATE:
                CTSdtDate date = sdtPr.addNewDate();
                Calendar calendar = Calendar.getInstance();

                CTString format = CTString.Factory.newInstance();
                format.setVal("yyyy/MM/dd hh:mm:ss");

                date.setDateFormat(format);
                date.setFullDate(calendar);

        }

        ctSdtRun.addNewSdtContent().addNewR().addNewT().setStringValue("请选择");

        run = para.createRun();
        run.setText("] 内容控件后");
    }

    /**
     * 创建页眉/页脚
     * @param header
     * @param footer
     */
    public void createHeaderFooter(String header, String footer) {
        createHeaderFooter(header, footer, true, true);
    }

    /**
     * 创建页眉/页脚
     * @param header
     * @param footer
     */
    public void createHeaderFooter(String header, String footer, boolean isPageSet, boolean isPageCountSet) {
        XWPFHeader hdr = document.createHeader(HeaderFooterType.DEFAULT);
        hdr.createParagraph().createRun().setText(header);

        XWPFFooter ftr = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph paraFooter = ftr.createParagraph();
        paraFooter.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paraFooter.createRun();
        run.setText(footer);

        if (isPageSet) {
            //当前页码
            CTFldChar fldChar = run.getCTR().addNewFldChar();
            fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));
            setTextStyle(run, STYLE_FOOTER);
            run = paraFooter.createRun();
            CTText ctText = run.getCTR().addNewInstrText();
            ctText.setStringValue("PAGE  \\* MERGEFORMAT");
            ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
            setTextStyle(run, STYLE_FOOTER);
            fldChar = run.getCTR().addNewFldChar();
            fldChar.setFldCharType(STFldCharType.Enum.forString("end"));

            //分隔符
            run = paraFooter.createRun();
            run.setText("/");

            //总页数
            if (isPageCountSet) {
                fldChar = run.getCTR().addNewFldChar();
                fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));
                setTextStyle(run, STYLE_FOOTER);
                run = paraFooter.createRun();
                ctText = run.getCTR().addNewInstrText();
                ctText.setStringValue("NUMPAGES  \\\\* MERGEFORMAT");
                ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
                setTextStyle(run, STYLE_FOOTER);
                fldChar = run.getCTR().addNewFldChar();
                fldChar.setFldCharType(STFldCharType.Enum.forString("end"));
            }
        }
    }

    /**
     * 创建自定义TOC域
     */
    public void createCustomTOC() {
//        CTSdtBlock block = document.getDocument().getBody().addNewSdt();
        CTSdtBlock block = document.getDocument().getBody().insertNewSdt(1);
        List<IBodyElement> list = document.getBodyElements();
        List<CTSdtBlock> list1 = document.getDocument().getBody().getSdtList();
        CustomTOC toc = new CustomTOC(block);
        Iterator i$ = document.getParagraphs().iterator();

        while(i$.hasNext()) {
            XWPFParagraph par = (XWPFParagraph)i$.next();
            String parStyle = par.getStyle();
            if (parStyle != null && parStyle.startsWith("Heading")) {
                try {
                    int level = Integer.parseInt(parStyle.substring("Heading".length()));
                    toc.addRow(level, par.getText(), 1, "112723803");
                } catch (NumberFormatException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    /**
     * 创建TOC域
     */
    public void createTOC() {
//        createCustomTOC();
        try {
            generateTOC();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成TOC域
     * @throws InvalidFormatException
     * @throws FileNotFoundException
     * @throws IOException
     */
    private void generateTOC() throws InvalidFormatException, FileNotFoundException, IOException {
        String findText = "{toc}";
        String replaceText = "";

        for (XWPFParagraph p : document.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                int pos = r.getTextPosition();
                String text = r.getText(pos);
//                System.out.println(text);
                if (text != null && text.contains(findText)) {
                    text = text.replace(findText, replaceText);
                    r.setText(text, 0);
                    addField(p, "TOC \\o \"1-3\" \\h \\z \\u");
//                    addField(p, "TOC \\h");
                    break;
                }
            }
        }
    }

    /**
     * 添加域
     * @param paragraph
     * @param fieldName
     */
    private void addField(XWPFParagraph paragraph, String fieldName) {
        CTSimpleField ctSimpleField = paragraph.getCTP().addNewFldSimple();
        ctSimpleField.setInstr(fieldName);
        ctSimpleField.setDirty(STOnOff.TRUE);
        ctSimpleField.addNewR().addNewT().setStringValue("<<fieldName>>");
    }

    /**
     * 替换文档中的指定文本
     * @param src
     * @param des
     */
    public void replaceText(String src, String des) {
        for (XWPFParagraph para : document.getParagraphs()) {
            //word有可能将一个Paragraph中的文本拆分为几个XWPFRun
            String strRun = "";

            for (XWPFRun run : para.getRuns()) {
                String str = run.getText(run.getTextPosition());
                if (null != str) {
                    strRun += str;
                }
            }

            if (strRun.contains(src)) {
                strRun = strRun.replace(src, des);
                removeRuns(para);

                XWPFRun run = para.insertNewRun(0);
                run.setText(strRun);
            }
        }
    }

    /**
     * 替换文档中的指定文本
     * @param map
     */
    public void replaceText(HashMap<String, String> map) {
        for (XWPFParagraph para : document.getParagraphs()) {
            //word有可能将一个Paragraph中的文本拆分为几个XWPFRun
            String strRun = "";
            TextStyle style = null;

            for (XWPFRun run : para.getRuns()) {
                String str = run.getText(run.getTextPosition());
                style = getTextStyle(run);

                if (null != str) {
                    strRun += str;
                }
            }

            for (Map.Entry<String, String> entry : map.entrySet()) {
                if (strRun.contains(entry.getKey())) {
                    strRun = strRun.replace(entry.getKey(), entry.getValue());

                    //因为一段文本可能由多个run组成，所以无法直接替换，只能先删除所有run，然后添加一个run
                    removeRuns(para);

                    XWPFRun run = para.insertNewRun(0);
                    run.setText(strRun);

                    if (null != style) {
                        setTextStyle(run, style);
                    }
                }
            }
        }
    }

    /**
     * 获取文本样式
     * @param run
     * @return
     */
    private TextStyle getTextStyle(XWPFRun run) {
        TextStyle textStyle = new TextStyle();
        textStyle.setFontFamily(run.getFontFamily());
        textStyle.setFontSize(String.valueOf(run.getFontSize()));
        textStyle.setColorVal(run.getColor());
        textStyle.setBold(run.isBold());
        return textStyle;
    }

    /**
     * 设置文本样式
     * @param run
     * @param textStyle
     */
    private void setTextStyle(XWPFRun run, TextStyle textStyle) {
        run.setFontFamily(textStyle.getFontFamily());
        run.setFontSize(Integer.parseInt(textStyle.getFontSize()));
        run.setColor(textStyle.getColorVal());
        run.setBold(textStyle.isBold());
    }

    /**
     * 删除段落中的所有run
     * @param para 指定段落
     */
    private void removeRuns(XWPFParagraph para) {
        int size = para.getRuns().size();
        for (int i = 0; i < size; i++) {
            para.removeRun(0);
        }
    }

    /**
     * 替换文本
     * @param para 段落
     * @param src  被替换文本
     * @param des  替换文本
     */
    public void replaceText(XWPFParagraph para, String src, String des) {
        for (XWPFRun run : para.getRuns()) {
            String str = run.getText(run.getTextPosition());
            str = str.replace(src, des);
            run.setText(str);
        }
    }

    /**
     * 设置纸张方向
     * @param orientation   枚举PAGE_ORIENTATION：HORIZONTAL 横向 / VERTICAL 纵向
     */
    private void setPageOrientation(PAGE_ORIENTATION orientation) {
        CTBody body = document.getDocument().getBody();

        if (!body.isSetSectPr()) {
            body.addNewSectPr();
        }
        CTSectPr section = body.getSectPr();

        if(!section.isSetPgSz()) {
            section.addNewPgSz();
        }
        CTPageSz pageSize = section.getPgSz();

        //必须要设置下面两个参数，否则整个的代码是无效的
        switch (orientation) {
            case HORIZONTAL:
                pageSize.setW(BigInteger.valueOf(PAGE_A4_HEIGHT));
                pageSize.setH(BigInteger.valueOf(PAGE_A4_WIDTH));
                break;
            case VERTICAL:
                pageSize.setW(BigInteger.valueOf(PAGE_A4_WIDTH));
                pageSize.setH(BigInteger.valueOf(PAGE_A4_HEIGHT));
                break;
        }
        pageSize.setOrient(STPageOrientation.LANDSCAPE);
    }


    /**
     * 添加分栏
     * @param para          在该段落上添加分栏
     * @param num           需要分成栏数
     * @throws XmlException
     */
    public void createColumn(XWPFParagraph para, int num) throws XmlException {
        createColumn(para, orientation, num);
    }

    /**
     * 添加分栏
     * @param para          在该段落上添加分栏
     * @param orientation   枚举PAGE_ORIENTATION：HORIZONTAL 横向 / VERTICAL 纵向
     * @param num           需要分成栏数
     * @throws XmlException
     */
    public void createColumn(XWPFParagraph para, PAGE_ORIENTATION orientation, int num) throws XmlException {
        CTSectPr section;

        if (null != para) {
            CTBody body = document.getDocument().getBody();
            section = body.getSectPr();

            CTColumns columns = section.getCols();
            columns.setNum(BigInteger.valueOf(num));
        } else {
//            section = para.getCTP().getPPr().getSectPr();
            para = document.createParagraph();
            section = para.getCTP().addNewPPr().addNewSectPr();
            section.addNewType().setVal(STSectionMark.CONTINUOUS);

            CTPageSz pgSz = section.addNewPgSz();

            switch (orientation) {
                case HORIZONTAL:
                    pgSz.setW(BigInteger.valueOf(PAGE_A4_HEIGHT));
                    pgSz.setH(BigInteger.valueOf(PAGE_A4_WIDTH));
                    break;
                case VERTICAL:
                    pgSz.setW(BigInteger.valueOf(PAGE_A4_WIDTH));
                    pgSz.setH(BigInteger.valueOf(PAGE_A4_HEIGHT));
                    break;
            }

//            CTPageMar pgMar = section.addNewPgMar();
//            pgMar.setTop(BigInteger.valueOf(1440));
//            pgMar.setRight(BigInteger.valueOf(1800));
//            pgMar.setBottom(BigInteger.valueOf(1440));
//            pgMar.setLeft(BigInteger.valueOf(1800));
//            pgMar.setHeader(BigInteger.valueOf(851));
//            pgMar.setFooter(BigInteger.valueOf(992));
//            pgMar.setGutter(BigInteger.valueOf(0));

            CTColumns columns = section.addNewCols();
            columns.setNum(BigInteger.valueOf(num));
//            columns.setSpace(BigInteger.valueOf(312));

//            CTDocGrid docGrid = section.addNewDocGrid();
//            docGrid.setLinePitch(BigInteger.valueOf(312));
////            docGrid.xsetType(new STDocGridImpl(new SchemaTypeImpl()));
//            docGrid.setType(STDocGrid.LINES);

        }
    }

    /**
     * 添加分节
     */
    @Deprecated
    public void createSection() {
        XWPFParagraph para = document.createParagraph();
        CTSectPr section = para.getCTP().addNewPPr().addNewSectPr();
        section.addNewType().setVal(STSectionMark.CONTINUOUS);

        CTPageSz pgSz = section.addNewPgSz();
        pgSz.setW(BigInteger.valueOf(PAGE_A4_HEIGHT));
        pgSz.setH(BigInteger.valueOf(PAGE_A4_WIDTH));

        CTColumns columns = section.addNewCols();
        columns.setNum(BigInteger.valueOf(1));
    }

    /**
     * 生成柱状图
     * @param titles
     * @param categories
     * @param values
     */
    public void createColumnChart(String[] titles, String[] categories, Double[][] values) {
        ColumnChart.createChart(document, titles, categories, values);
    }

    /**
     * 生成折线图
     * @param title
     * @param xAxisTitle
     * @param yAxisTitle
     * @param categories
     * @param values
     */
    public void createLineChart(String title, String xAxisTitle, String[] yAxisTitle, String[] categories, Double[][] values) {
        LineChart.createChart(document, title, xAxisTitle, yAxisTitle, categories, values);
    }

    /**
     * 生成饼图
     * @param title
     * @param categories
     * @param values
     */
    public void createPieChart(String title, String[] categories, Double[] values) {
        PieChart.createChart(document, title, categories, values);
    }

    /**
     * 保存文档
     * @throws IOException
     */
    public void save() throws IOException {
//        document.enforceUpdateFields();
//        System.out.println(document.isEnforcedUpdateFields());
        xwpfHelper.saveDocument(document, fileName);
        //测试：获取文档所有样式
//        getStyleValue();
//        PythonUtils.updateToc(fileName);

        // 使用jacob自动更新文档目录并保存文档
        // ----------------------重要信息---------------------
        // 此处如报错，请查看readme.md文件，将对应的库文件放入jdk目录下
        // ----------------------重要信息---------------------
        Dispatch doc = jacobOpenDoc();
        jacobCloseDoc(doc);
    }

    /**
     * 使用jacob方式打开word文档
     * @return  打开的文档对象
     */
    private Dispatch jacobOpenDoc() {
//        ComThread.InitSTA();
        // 打开Word应用程序
        ActiveXComponent word = new ActiveXComponent("Word.Application");
        // 设置word不可见
        word.setProperty("Visible", new Variant(false));

        // 打开word文件
        Dispatch documents = word.getProperty("Documents").toDispatch();
        Dispatch doc = Dispatch.call(documents, "Open", fileName).toDispatch();

        return doc;
    }

    /**
     * 使用jacob方式关闭文档
     * @param doc   word文档对象
     */
    private void jacobCloseDoc(Dispatch doc) {
        if (doc != null) {
            //关闭文档且保存
//            Dispatch.call(doc, "Save");
            Dispatch.call(doc, "Close", new Variant(true));
            doc = null;
        }
//        ComThread.Release();
    }
}
