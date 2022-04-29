package com.quanxi.qxexportutils.util.doc.poi.word2html;

import fr.opensagres.poi.xwpf.converter.xhtml.Base64EmbedImgManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class ExportHtml {
    public static String doc2Html(String wordPath, String wordName, String suffix) throws Exception {
        ZipSecureFile.setMinInflateRatio(-1.0d);
        String htmlPath = wordPath + File.separator + "html"
                + File.separator;
        String htmlName = wordName + ".html";
        String imagePath = htmlPath + "image" + File.separator;

        // 判断html文件是否存在
        File htmlFile = new File(htmlPath + htmlName);
//      if (htmlFile.exists()) {
//          return htmlFile.getAbsolutePath();
//      }

        // word文件
        File wordFile = new File(wordPath + File.separator + wordName + suffix);

        // 1) 加载word文档生成 XWPFDocument对象
        InputStream in = new FileInputStream(wordFile);
        XWPFDocument document = new XWPFDocument(in);

        // 2) 解析 XHTML配置 (这里设置IURIResolver来设置图片存放的目录)
        File imgFolder = new File(imagePath);
        //  带图片的word，则将图片转为base64编码，保存在一个页面中
        XHTMLOptions options = XHTMLOptions.create().indent(4).setImageManager(new Base64EmbedImgManager());
        // 3) 将 XWPFDocument转换成XHTML
        // 生成html文件上级文件夹
        File folder = new File(htmlPath);
        if (!folder.exists()) {
            folder.mkdirs();
        }
        OutputStream out = new FileOutputStream(htmlFile);
        XHTMLConverter.getInstance().convert(document, out, options);

        return htmlFile.getAbsolutePath();
    }
}
