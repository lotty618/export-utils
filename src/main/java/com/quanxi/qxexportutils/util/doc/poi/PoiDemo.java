package com.quanxi.qxexportutils.util.doc.poi;

import com.quanxi.qxexportutils.common.DocConst;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.HashMap;
import java.util.Map;

public class PoiDemo {
    public static void main(String[] args) {
        try {
            //新建Word工具对象
//            PoiWordUtils wordUtils = new PoiWordUtils(System.getProperty("user.dir") + "\\qcs_doc_cover.docx", System.getProperty("user.dir") + "\\testpoi.docx", PoiWordUtils.PAGE_ORIENTATION.HORIZONTAL);
            PoiWordUtils wordUtils = new PoiWordUtils(System.getProperty("user.dir") + "\\testpoi20.docx", PoiWordUtils.PAGE_ORIENTATION.HORIZONTAL);
            //目录是否分栏显示
            boolean isSeparateColumns = true;

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

            //添加新页段落
            wordUtils.createParagrah("1. 这里是标题", true, PoiWordUtils.STYLE_TITLE, true);

            //替换文本
            HashMap<String, String> map = new HashMap<>();
            map.put("${hos}", "龙岗中心医院");
            map.put("${report}", "2019年第四季度");
            wordUtils.replaceText(map);

            //添加表格
            for (int i = 0; i < 10; i++) {

                //表格标题
                wordUtils.createParagrah("新页面新标题" + i, true, PoiWordUtils.STYLE_TITLE, true);

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
                long[] widths = {1500, 2000*5, 1500, 1500, 1500};
                //生成表格
                XWPFTable table = wordUtils.createTable(header, PoiWordUtils.STYLE_TABLE_SUB_HEADER, data, PoiWordUtils.STYLE_TABLE_SUB_CONTENT, widths, 25, new int[]{0}, new int[]{4});
            }

            //////////////
            Map<String, String> mpData = new HashMap<>();
            mpData.put("0", "苹果");
            mpData.put("1", "香蕉");
            wordUtils.createControl(DocConst.CONTROL_TYPE.DROPDOWNLIST, mpData);

            wordUtils.createControl(DocConst.CONTROL_TYPE.COMBOBOX, mpData);

            wordUtils.createControl(DocConst.CONTROL_TYPE.DATE, null);

            //创建柱状图
            //标题
            wordUtils.createParagrah("各国数据对比1", true, PoiWordUtils.STYLE_TITLE, true);
            // create the data
            String[] titles = new String[] {"GDP", "CPI"};
            String[] categories = new String[] { "中国", "美国", "俄罗斯", "日本" };
            Double[][] values = {
                    new Double[] { 100d, 120d, 90d, 95d },
                    new Double[] { 4.2d, 3.5d, 4.3d, 3.2d }
            };
            wordUtils.createColumnChart(titles, categories, values);

            //页眉页脚
//            wordUtils.createHeaderFooter("这是页眉", "这是页脚");

            //生成目录
            wordUtils.createTOC();

            //保存Word
            wordUtils.save();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
