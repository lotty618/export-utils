package com.quanxi.qxexportutils.util.doc.poi;

import com.quanxi.qxexportutils.util.doc.poi.export2word.XWPFHelperTable;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class DemoEx {
    public static void main(String[] args) {
        try {
            PoiWordUtils wordUtils = new PoiWordUtils(System.getProperty("user.dir") + "\\test1.docx", PoiWordUtils.PAGE_ORIENTATION.HORIZONTAL);
//            String[] header = new String[]{"", "", "", "", ""};
//            String[][] data = new String[][]{
//                    header,
//                    header,
//                    header,
//                    header,
//                    header,
//            };
//            long[] widths = {1500, 1500, 1500, 1500, 1500};
//            XWPFTable table = wordUtils.createTable(header, PoiWordUtils.STYLE_TABLE_SUB_HEADER, data, PoiWordUtils.STYLE_TABLE_SUB_CONTENT, widths, 25, null, null);
//            XWPFHelperTable helperTable = new XWPFHelperTable();
//            helperTable.mergeCellsHorizontal(table, 5, 3, 4);
//目录是否分栏显示
            boolean isSeparateColumns = false;

            //添加目录标题
            wordUtils.createParagrah("临床医技科室2019年第四季度全面质量考核结果反馈目录", true, PoiWordUtils.STYLE_TITLE, false);
            //此为防止下面目录分栏使目录以上所有文档部分都分栏，则此处添加section
            if (isSeparateColumns) {
//                wordUtils.createSection();
                wordUtils.createColumn(null, 1);
            }
            //添加目录
            XWPFParagraph para = wordUtils.createParagrah("{toc}", false, PoiWordUtils.STYLE_TOC, false);

            //将目录分成两栏显示
            if (isSeparateColumns) {
                wordUtils.createColumn(null, 2);
            }

            //表格标题
            wordUtils.createParagrah("新页面新标题", true, PoiWordUtils.STYLE_TITLE, true);

            //表格标题
            String[] header = new String[]{"督导项目", "存在问题/奖惩/亮点", "扣分", "奖励(元)", "处罚(元)"};
            //表格数据
            String[][] data = new String[][]{
                    new String[]{"医务管理", "1.住院号464250，（1）《手术风险评估表》手术医生未评估签字；（2）《患者手术知情同意书》未执行双签名，手术医生未签名。-0.1\n" +
                            "2.抽查病历住院号：303442，医生手写签名字迹不一致。抽查住院号464321病历，无授权委托书。-0.05\n" +
                            "3.住院号457230，《手术安全核查》医师未签名。-0.05", "0.2", "", ""},
                    new String[]{"信息管理", "病历回收绩效：（每月统计各临床科室病历7个工作日回收率，要求回收率达100%，每低于1%，扣0.1分，扣完为止。）11月份7个工作日回收率61.5%，不达标（7个工作日病历回收率要求100%），扣0.6分；11月份7个工作日回收率70.43%，不达标，扣0.4分。", "1", "", ""},
                    new String[]{"信息管理", "其他项，不达标，扣0.2分。", "0.5", "", ""},
                    new String[]{"其他项目", "其他项，不达标，扣0.3分。", "0.3", "", ""},
            };
            //表格宽度
            long[] widths = new long[]{1500, 2000*5, 1500, 1500, 1500};
            //生成表格
            XWPFTable table = wordUtils.createTable(header, PoiWordUtils.STYLE_TABLE_SUB_HEADER, data, PoiWordUtils.STYLE_TABLE_SUB_CONTENT, widths, 25, new int[]{0}, new int[]{4});
//            XWPFHelperTable helperTable = new XWPFHelperTable();
////                helperTable.mergeCellsVertically(table, 4, 3, 4);
//            helperTable.mergeCellsHorizontal(table, 4, 3, 4);

            //生成目录
            wordUtils.createTOC();
            wordUtils.save();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
