package com.quanxi.qxexportutils.util.doc.poi.export2word;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExportWord {
    private XWPFHelperTable xwpfHelperTable = null;
    private XWPFHelper xwpfHelper = null;
    public ExportWord() {
        xwpfHelperTable = new XWPFHelperTable();
        xwpfHelper = new XWPFHelper();
    }
    /**
     * 创建好文档的基本 标题，表格  段落等部分
     * @return
     */
    public XWPFDocument createXWPFDocument() {
        XWPFDocument doc = new XWPFDocument();
        createTitleParagraph(doc, false);
        createTextParagraph(doc, false, "1.复杂性高\\n\" +\n" +
                "                \"整个项目包含的模块非常多，模块的边界模糊，依赖关系不清晰，代码质量参差不齐,整个项目非常复杂。每次修改代码都心惊胆战，甚至添加一个简单的功能，或者修改一个BUG都会造成隐含的缺陷。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"2.技术债务逐渐上升\\n\" +\n" +
                "                \"随着时间推移、需求变更和人员更迭，会逐渐形成应用程序的技术债务，并且越积越多。已使用的系统设计或代码难以修改，因为应用程序的其他模块可能会以意料之外的方式使用它。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"3.部署速度逐渐变慢，频率低\\n\" +\n" +
                "                \"随着代码的增加，构建和部署的时间也会增加。而在单体应用中，每次功能的变更或缺陷的修复都会导致我们需要重新部署整个应用。全量部署的方式耗时长、影响范围大、风险高，这使得单体应用项目上线部署的频率较低，从而又导致两次发布之间会有大量功能变更和缺陷修复，出错概率较高。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"4.扩展能力受限，无法按需伸缩\\n\" +\n" +
                "                \"单体应用只能作为一个整体进行扩展，无法结合业务模块的特点进行伸缩。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"5.阻碍技术创新\\n\" +\n" +
                "                \"单体应用往往使用统一的技术平台或方案解决所有问题，团队的每个成员都必须使用相同的开发语言和架构，想要引入新的框架或技术平台非常困难。1.复杂性高\\n\" +\n" +
                "                \"整个项目包含的模块非常多，模块的边界模糊，依赖关系不清晰，代码质量参差不齐,整个项目非常复杂。每次修改代码都心惊胆战，甚至添加一个简单的功能，或者修改一个BUG都会造成隐含的缺陷。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"2.技术债务逐渐上升\\n\" +\n" +
                "                \"随着时间推移、需求变更和人员更迭，会逐渐形成应用程序的技术债务，并且越积越多。已使用的系统设计或代码难以修改，因为应用程序的其他模块可能会以意料之外的方式使用它。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"3.部署速度逐渐变慢，频率低\\n\" +\n" +
                "                \"随着代码的增加，构建和部署的时间也会增加。而在单体应用中，每次功能的变更或缺陷的修复都会导致我们需要重新部署整个应用。全量部署的方式耗时长、影响范围大、风险高，这使得单体应用项目上线部署的频率较低，从而又导致两次发布之间会有大量功能变更和缺陷修复，出错概率较高。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"4.扩展能力受限，无法按需伸缩\\n\" +\n" +
                "                \"单体应用只能作为一个整体进行扩展，无法结合业务模块的特点进行伸缩。\\n\" +\n" +
                "                \"\\n\" +\n" +
                "                \"5.阻碍技术创新\\n\" +\n" +
                "                \"单体应用往往使用统一的技术平台或方案解决所有问题，团队的每个成员都必须使用相同的开发语言和架构，想要引入新的框架或技术平台非常困难。1.复杂性高\\n\" +\n" +
                "                \"整个项目包含的模块非常多，模块的边界模糊，依赖关系不清晰，代码质量参差不齐,整个项目非常复杂。每次修改代码都心惊胆战，甚至添加一个简单的功能，或者修改一个BUG都会造成隐含的缺陷。\\n\" +\n" +
                "                \"\\n\" +\n");
        createTableParagraph(doc, true, 10, 6);

        return doc;
    }
    /**
     * 创建表格的标题样式
     * @param document
     */
    public void createTitleParagraph(XWPFDocument document, boolean newPage) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        //是否创建新页
        titleParagraph.setPageBreak(newPage);
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);//样式居中
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
//        titleFun.setText(titleName); //设置标题的名字
        titleFun.setBold(true); //加粗
        titleFun.setColor("000000");//设置颜色
        titleFun.setFontSize(25);    //字体大小
//        titleFun.setFontFamily("");//设置字体
        //...
        titleFun.addBreak();    //换行

    }

    /**
     *
     * @param document
     */
    public void createTextParagraph(XWPFDocument document, boolean newPage, String text) {
        XWPFParagraph textParagraph = document.createParagraph();    //新建一个段落对象（就是一段文字）
        //是否创建新页
        textParagraph.setPageBreak(newPage);
        textParagraph.setAlignment(ParagraphAlignment.LEFT);//样式居左
        XWPFRun titleFun = textParagraph.createRun();    //创建文本对象
        titleFun.setText(text); //设置标题的名字
        titleFun.setColor("000000");//设置颜色
        titleFun.setFontSize(10);    //字体大小
//        titleFun.setFontFamily("");//设置字体
        //...
        titleFun.addBreak();    //换行
    }

    /**
     * 设置表格
     * @param document
     * @param rows
     * @param cols
     */
    public void createTableParagraph(XWPFDocument document, boolean newPage, int rows, int cols) {
//        xwpfHelperTable.createTable(xdoc, rowSize, cellSize, isSetColWidth, colWidths)
        XWPFParagraph textParagraph = document.createParagraph();    //新建一个段落对象（就是一段文字）
        //是否创建新页
        //textParagraph.setPageBreak(newPage);
        textParagraph.createRun().addBreak(BreakType.PAGE);

        XWPFTable infoTable = document.createTable(rows, cols);
        xwpfHelperTable.setTableWidthAndHAlign(infoTable, "9072", STJc.CENTER);
        //合并表格
        xwpfHelperTable.mergeCellsHorizontal(infoTable, 1, 1, 5);
        xwpfHelperTable.mergeCellsVertically(infoTable, 0, 3, 6);
        for(int col = 3; col < 7; col++) {
            xwpfHelperTable.mergeCellsHorizontal(infoTable, col, 1, 5);
        }
        //设置表格样式
        List<XWPFTableRow> rowList = infoTable.getRows();
        for(int i = 0; i < rowList.size(); i++) {
            XWPFTableRow infoTableRow = rowList.get(i);
            List<XWPFTableCell> cellList = infoTableRow.getTableCells();
            for(int j = 0; j < cellList.size(); j++) {
                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun cellParagraphRun = cellParagraph.createRun();
                cellParagraphRun.setFontSize(12);
                if(i % 2 != 0) {
                    cellParagraphRun.setBold(true);
                }
            }
        }
        xwpfHelperTable.setTableHeight(infoTable, 560, STVerticalJc.CENTER);
    }

    /**
     * 往表格中填充数据
     * @param dataList
     * @param document
     * @throws IOException
     */
    @SuppressWarnings("unchecked")
    public void exportCheckWord(Map<String, Object> dataList, XWPFDocument document, String savePath) throws IOException {
        XWPFParagraph paragraph = document.getParagraphArray(0);
        XWPFRun titleFun = paragraph.getRuns().get(0);
        titleFun.setText(String.valueOf(dataList.get("TITLE")));
        List<List<Object>> tableData = (List<List<Object>>) dataList.get("TABLEDATA");
        XWPFTable table = document.getTableArray(0);
        fillTableData(table, tableData);
        xwpfHelper.saveDocument(document, savePath);
    }
    /**
     * 往表格中填充数据
     * @param table
     * @param tableData
     */
    public void fillTableData(XWPFTable table, List<List<Object>> tableData) {
        List<XWPFTableRow> rowList = table.getRows();
        for(int i = 0; i < tableData.size(); i++) {
            List<Object> list = tableData.get(i);
            List<XWPFTableCell> cellList = rowList.get(i).getTableCells();
            for(int j = 0; j < list.size(); j++) {
                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                XWPFRun cellParagraphRun = cellParagraph.getRuns().get(0);
                cellParagraphRun.setText(String.valueOf(list.get(j)));
            }
        }
    }
}
