package com.quanxi.qxexportutils.util.doc.poi;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Test {
    public static void main(String[] args) {
        try {
            FileInputStream is = new FileInputStream("D:/test1.docx");
            OPCPackage open = OPCPackage.open(is);
            List<XWPFAbstractSDT> sdts = new ArrayList<>();

            XWPFDocument document = new XWPFDocument(open);
            List<XWPFParagraph> paragraphs = document.getParagraphs();

            for (XWPFParagraph paragraph : paragraphs) {

                List<CTSdtRun> sdtList = paragraph.getCTP().getSdtList();
                for (CTSdtRun ctSdtRun : sdtList) {
                    CTSdtPr sdtPr = ctSdtRun.getSdtPr();


                    List<CTRPr> rPrList = sdtPr.getRPrList();

                    List<CTSdtDate> dateList = sdtPr.getDateList();
                    for (CTSdtDate ctSdtDate : dateList) {
                        System.out.println("===== Date: " + ctSdtDate.getFullDate().toString());
                    }

                    List<CTSdtDropDownList> dropDownListList = sdtPr.getDropDownListList();
                    for (int i = 0; i < dropDownListList.size(); i++) {
                        CTSdtDropDownList dropDownList = dropDownListList.get(i);

                        System.out.println("===== DropDownList: ");

                        for (CTSdtListItem ctSdtListItem : dropDownList.getListItemList()) {
                            System.out.println(ctSdtListItem.getDisplayText() + " : " + ctSdtListItem.getValue());
                        }
                    }

//                    CTSdtContentRun sdtContent = ctSdtRun.getSdtContent();
//                    CTSdtEndPr sdtEndPr = ctSdtRun.getSdtEndPr();
//                    Node domNode = ctSdtRun.getDomNode();

                }

                List<IRunElement> iRuns = paragraph.getIRuns();
                for (IRunElement iRun : iRuns) {
                    if (iRun instanceof  XWPFSDT) {
                        sdts.add((XWPFSDT)iRun);
                    }
                }
            }

            List<POIXMLDocumentPart> relations = document.getRelations();
            System.out.println(relations.size());

        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }
//
//    private List<XWPFAbstractSDT> extractAllSDTs(XWPFDocument doc) {
//        List<XWPFAbstractSDT> sdts = new ArrayList<>();
//
//        doc.getRelations()
//
//        List<XWPFHeader> headers = doc.getHeaderList();
//        for (XWPFHeader header : headers) {
//            sdts.addAll(extractSDTsFromBodyElements(header.getBodyElements()));
//        }
//        sdts.addAll(extractSDTsFromBodyElements(doc.getBodyElements()));
//
//        List<XWPFFooter> footers = doc.getFooterList();
//        for (XWPFFooter footer : footers) {
//            sdts.addAll(extractSDTsFromBodyElements(footer.getBodyElements()));
//        }
//
//        for (XWPFFootnote footnote : doc.getFootnotes()) {
//            sdts.addAll(extractSDTsFromBodyElements(footnote.getBodyElements()));
//        }
//        for (XWPFEndnote footnote : doc.getEndnotes()) {
//            sdts.addAll(extractSDTsFromBodyElements(footnote.getBodyElements()));
//        }
//        return sdts;
//    }
}
