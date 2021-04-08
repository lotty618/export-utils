package com.quanxi.qxexportutils.util.doc.poi.export2word.chart;

import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import sun.awt.SunHints;

import java.io.FileOutputStream;

/**
 * 饼图
 */
public class PieChart {

    /**
     * 生成饼图
     * @param document      Word文档
     * @param title         标题
     * @param categories    类型
     * @param values        数据
     */
    public static void createChart(XWPFDocument document, String title, String[] categories, Double[] values) {
        try {
            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

            //标题
            chart.setTitleText(title);
            //标题是否覆盖图表
            chart.setTitleOverlay(false);

            //图例位置
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);

            //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
            //分类轴标数据，
//			XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
            XDDFCategoryDataSource cData = XDDFDataSourcesFactory.fromArray(categories);
            //数据1，
//			XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
            XDDFNumericalDataSource<Double> vData = XDDFDataSourcesFactory.fromArray(values);
            //XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
            XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
            //设置为可变颜色
            data.setVaryColors(true);
            //图表加载数据
            data.addSeries(cData, vData);

            //绘制
            chart.plot(data);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws Exception {
        try (XWPFDocument document = new XWPFDocument()) {

            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

            //标题
            chart.setTitleText("地区排名前七的国家");
            //标题是否覆盖图表
            chart.setTitleOverlay(false);

            //图例位置
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);

            //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
            //分类轴标数据，
//			XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
            XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(new String[] {"俄罗斯","加拿大","美国","中国","巴西","澳大利亚","印度"});
            //数据1，
//			XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
            XDDFNumericalDataSource<Integer> values = XDDFDataSourcesFactory.fromArray(new Integer[] {17098242,9984670,9826675,9596961,8514877,7741220,3287263});
            //XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
            XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
            //设置为可变颜色
            data.setVaryColors(true);
            //图表加载数据
            data.addSeries(countries, values);


            //绘制
            chart.plot(data);

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
                document.write(fileOut);
            }
        }
    }
}
