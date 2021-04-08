package com.quanxi.qxexportutils.util.doc.spiredoc;

import com.spire.doc.*;
import com.spire.doc.documents.*;

public class CreateWordDocument {
    public static void main(String[] args) {
        //合并文档
        WordUtils wordUtils = new WordUtils("qcs_doc_cover.docx", "test1.docx");

        //添加目录
        Section sectionContents = wordUtils.createSection();
        String contents = "临床医技科室2019年第四季度全面质量考核结果反馈目录";
        Paragraph paraContents = wordUtils.insertParagraph(sectionContents, contents, WordUtils.STYLE_TITLE);
        paraContents.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//
//        paraContents.insertSectionBreak(SectionBreakType.No_Break);

        //目录一栏
//        Paragraph para = sectionContents.addParagraph();
//        wordUtils.addTableOfContents(para);

        //目录分栏
//        Section sectionc = wordUtils.createSection();
//        //默认创建的Section自带分页符，如不需要则设置为SectionBreakType.No_Break
//        sectionc.setBreakCode(SectionBreakType.No_Break);
//
//        sectionc.addColumn(75f, 20f);
//        sectionc.addColumn(75f, 20f);
//        Paragraph para = sectionc.addParagraph();
        wordUtils.addTableOfContents(paraContents);

        //替换文本
        wordUtils.replaceText("${city}", "XX市");
        wordUtils.replaceText("${hos_name}", "XX中心医院\n第X人民医院");
        wordUtils.replaceText("${hos}", "XX中心医院");
        wordUtils.replaceText("${title}", "临床医技科室全面质量考核结果反馈");
        wordUtils.replaceText("${report}", "2020年第四季度");

        //添加总表格
        Section section = wordUtils.createSection();
        Paragraph para = wordUtils.insertParagraph(section, "临床医技科室2019年第四季度全面质量考核结果汇总表", BuiltinStyle.Heading_1, WordUtils.STYLE_TITLE_HEADER);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        String[] header = {"科室", "总得分", "奖励(元)", "处罚(元)", "科室", "总得分", "奖励(元)", "处罚(元)", "科室", "总得分", "奖励(元)", "处罚(元)"};
        String[][] data = {
                new String[]{"普外科", "98.67", "15500", "274"},
        };
        wordUtils.insertTable(section, header, data, WordUtils.STYLE_SUM_TABLE_TITLE, WordUtils.STYLE_SUM_TABLE_CONTENT);
        para = wordUtils.insertParagraph(section, "说明：1、奖励项目：医务管理（住院总、医疗保障、单病种、不良事件）；护理管理（竞赛、硕士生导师、不良事件、继" +
                "续班）；临床医学院（科研、论文、继教班）；信息管理（病历回收率达标）；质控管理（荣誉项目、6S内审员）" +
                "健教管理（新闻宣传）；输血管理（自体输血）。", WordUtils.STYLE_PARA);
        para = wordUtils.insertParagraph(section, "2、处罚项目：药事管理（抗菌药物、I类切口专项点评）；信息管理（病历逾期未交）；物价管理（违规收费）。" +
                "医院质量与安全管理委员会", WordUtils.STYLE_PARA_BELOW);

        para = wordUtils.insertParagraph(section, "医院质量与安全管理委员会\n" +
                "2020年1月31日\n", WordUtils.STYLE_PARA);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
        para.getFormat().setRightIndent(32f);

        para = wordUtils.insertParagraph(section, "院长：                分管领导：                            质控科科长：                 制表人：             \n" +
                "分管财务院领导：                         财务科科长：                  ", WordUtils.STYLE_PARA);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //添加子表格
        section = wordUtils.createSection();
        para = wordUtils.insertParagraph(section, "普通外科2019年第四季度全面质量考核结果汇总表", BuiltinStyle.Heading_1, WordUtils.STYLE_TITLE_HEADER);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        header = new String[]{"督导项目", "存在问题/奖惩/亮点", "扣分", "奖励(元)", "处罚(元)"};
        data = new String[][]{
                new String[]{"医务管理", "1.住院号464250，（1）《手术风险评估表》手术医生未评估签字；（2）《患者手术知情同意书》未执行双签名，手术医生未签名。-0.1\n" +
                        "2.抽查病历住院号：303442，医生手写签名字迹不一致。抽查住院号464321病历，无授权委托书。-0.05\n" +
                        "3.住院号457230，《手术安全核查》医师未签名。-0.05", "0.2", "", ""},
                new String[]{"信息管理", "病历回收绩效：（每月统计各临床科室病历7个工作日回收率，要求回收率达100%，每低于1%，扣0.1分，扣完为止。）11月份7个工作日回收率61.5%，不达标（7个工作日病历回收率要求100%），扣0.6分；11月份7个工作日回收率70.43%，不达标，扣0.4分。", "1", "", ""},
        };
        float[] width = new float[]{20, 400, 15, 20, 20};
        float height = 25;
        wordUtils.insertTable(section, width, height, header, data, WordUtils.STYLE_SUB_TABLE_TITLE, WordUtils.STYLE_SUB_TABLE_CONTENT);

        for (int i = 0; i < 50; i++) {
            //添加标题
            Section section1 = wordUtils.createSection();
            Paragraph para1 = wordUtils.insertParagraph(section1, "测试科室2019年第四季度全面质量考核结果汇总表" + i, BuiltinStyle.Heading_1, WordUtils.STYLE_TITLE_HEADER);
            para1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

            //添加子表格
            header = new String[]{"督导项目", "存在问题/奖惩/亮点", "扣分", "奖励(元)", "处罚(元)"};
            data = new String[][]{
                    new String[]{"医务管理", "1.住院号464250，（1）《手术风险评估表》手术医生未评估签字；（2）《患者手术知情同意书》未执行双签名，手术医生未签名。-0.1\n" +
                            "2.抽查病历住院号：303442，医生手写签名字迹不一致。抽查住院号464321病历，无授权委托书。-0.05\n" +
                            "3.住院号457230，《手术安全核查》医师未签名。-0.05", "0.2", "", ""},
                    new String[]{"信息管理", "病历回收绩效：（每月统计各临床科室病历7个工作日回收率，要求回收率达100%，每低于1%，扣0.1分，扣完为止。）11月份7个工作日回收率61.5%，不达标（7个工作日病历回收率要求100%），扣0.6分；11月份7个工作日回收率70.43%，不达标，扣0.4分。", "1", "", ""},
            };
            wordUtils.insertTable(section1, width, height, header, data, WordUtils.STYLE_SUB_TABLE_TITLE, WordUtils.STYLE_SUB_TABLE_CONTENT);

        }

        //更新目录
        wordUtils.updateTableOfContents();

        //保存文档
        wordUtils.save();

    }
}
