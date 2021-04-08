package com.quanxi.qxexportutils.util.doc.poi.export2word.chart;

import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.print.DocFlavor;
import java.io.FileOutputStream;
import java.util.List;

/**
 * 折线图
 */
public class LineChart {

    /**
     * 生成折线图
     * @param document      Word文档
     * @param title         总标题
     * @param xAxisTitle    X轴标题
     * @param yAxisTitle    Y轴标题：可有左/右标题
     * @param categories    X轴数据
     * @param values        Y轴数据：可有左/右数据对应Y轴左/右标题
     */
    public static void createChart(XWPFDocument document, String title, String xAxisTitle, String[] yAxisTitle, String[] categories, Double[][] values) {
        try {
            // create the chart
            XWPFChart chart = document.createChart(26 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

            //标题
            chart.setTitleText(title);
            //标题覆盖
            chart.setTitleOverlay(false);

            //图例位置
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP);

            //分类轴标(X轴),标题位置
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle(xAxisTitle);

            //值(Y轴)轴,标题位置
            XDDFValueAxis[] yAxis = new XDDFValueAxis[yAxisTitle.length];

            for (int i = 0; i < yAxisTitle.length; i++) {
                if (i == 0) {
                    yAxis[i] = chart.createValueAxis(AxisPosition.LEFT);
                } else {
                    yAxis[i] = chart.createValueAxis(AxisPosition.RIGHT);
                }
                yAxis[i].setTitle((yAxisTitle[i]));
            }

            //填充X轴数据
            XDDFCategoryDataSource xDataSource = XDDFDataSourcesFactory.fromArray(categories);
            //填充Y轴标题
            XDDFNumericalDataSource<Double>[] yDataSource = new XDDFNumericalDataSource[yAxisTitle.length];

            //填充Y轴数据
            for (int i = 0; i < yAxisTitle.length; i++) {
                yDataSource[i] = XDDFDataSourcesFactory.fromArray(values[i]);

                //定义折线图
                XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, yAxis[i]);

                //图表加载数据，折线1
                XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(xDataSource, yDataSource[i]);
                //折线图例标题
                series1.setTitle(yAxisTitle[i], null);
                //是否平滑：曲线/直线
                series1.setSmooth(false);
                //设置标记大小
                series1.setMarkerSize((short) 6);
                //设置标记样式，星星
                series1.setMarkerStyle(MarkerStyle.SQUARE);

                //绘制
                chart.plot(data);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws Exception {
        try (XWPFDocument document = new XWPFDocument()) {
            String title = "各国数据对比2";
            String xAxisTitle = "国家";
            String[] yAxisTitle = new String[] {"面积", "人口"};
            String[] categories = new String[] { "中国", "美国", "俄罗斯", "日本" };
            Double[][] values = new Double[][]{
                    new Double[] { 9600000d, 9370000d, 17098200d, 377962d },
                    new Double[] { 140005d, 33000d, 14600d, 12600d }
            };

            createChart(document, title, xAxisTitle, yAxisTitle, categories, values);

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
                document.write(fileOut);
            }
        }
    }
}
