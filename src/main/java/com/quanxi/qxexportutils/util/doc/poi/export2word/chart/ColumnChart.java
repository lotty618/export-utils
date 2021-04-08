package com.quanxi.qxexportutils.util.doc.poi.export2word.chart;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;

/**
 * 柱状图
 */
public class ColumnChart {
    // Methode to set title in the data sheet without creating a Table but using the sheet data only.
    // Creating a Table is not really necessary.
    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null) {
            row = sheet.createRow(0);
        }
        XSSFCell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    /**
     * 生成柱状图
     * @param document      Word文档
     * @param titles        Y轴类型
     * @param categories    X轴数据
     * @param values        Y轴数据
     */
    public static void createChart(XWPFDocument document, String[] titles, String[] categories, Double[][] values) {
        try {
            // create the chart
            XWPFChart chart = document.createChart(24 * Units.EMU_PER_CENTIMETER, 8 * Units.EMU_PER_CENTIMETER);

            // create data sources
            int numOfPoints = categories.length;
            String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
//            String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
//            String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
            XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
//            XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);
//            XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);

            String[] valuesDataRange = new String[titles.length];
            XDDFNumericalDataSource<Double>[] valuesData = new XDDFNumericalDataSource[titles.length];
            for (int i = 0; i < titles.length; i++) {
                valuesDataRange[i] = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
                valuesData[i] = XDDFDataSourcesFactory.fromArray(values[i], valuesDataRange[i], i + 1);
            }

            // create axis
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
            // Else first and last category is exactly on cross points and the bars are only half visible.
            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

            // create chart data
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

            // create series
            /*
                // if only one series do not vary colors for each bar
                ((XDDFBarChartData) data).setVaryColors(false);
                XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
                // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
                // Excel cannot opening the workbook after creatingg that incomplete Table.
                // So updating the chart data in Word is not possible.
                //series.setTitle("a", chart.setSheetTitle("a", 1));
                series.setTitle(title, setTitleInDataSheet(chart, title, 1));
            */

            /*
			   // if more than one series do vary colors of the series
			   ((XDDFBarChartData)data).setVaryColors(true);
			   series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
             */

            XDDFChartData.Series series = null;
            for (int i = 0; i < titles.length; i++) {
                series = data.addSeries(categoriesData, valuesData[i]);
                series.setTitle(titles[i], setTitleInDataSheet(chart, titles[i], i + 1));

            }

            // plot chart data
            chart.plot(data);

            // create legend
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.LEFT);
            legend.setOverlay(false);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws Exception {
        try (XWPFDocument document = new XWPFDocument()) {

            // create the data
            String[] categories = new String[] { "Lang 1", "Lang 2", "Lang 3" };
            Double[] valuesA = new Double[] { 10d, 20d, 30d };
            Double[] valuesB = new Double[] { 15d, 25d, 35d };

            // create the chart
            XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

            // create data sources
            int numOfPoints = categories.length;
            String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
            String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
            String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
            XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
            XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);
            XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);

            // create axis
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
            // Else first and last category is exactly on cross points and the bars are only half visible.
            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

            // create chart data
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

            // create series
            // if only one series do not vary colors for each bar
            ((XDDFBarChartData) data).setVaryColors(false);
            XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
            // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
            // Excel cannot opening the workbook after creatingg that incomplete Table.
            // So updating the chart data in Word is not possible.
            //series.setTitle("a", chart.setSheetTitle("a", 1));
            series.setTitle("a", setTitleInDataSheet(chart, "a", 1));

			/*
			   // if more than one series do vary colors of the series
			   ((XDDFBarChartData)data).setVaryColors(true);
			   series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
			*/

            // plot chart data
            chart.plot(data);

            // create legend
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.LEFT);
            legend.setOverlay(false);

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
                document.write(fileOut);
            }
        }
    }
}
