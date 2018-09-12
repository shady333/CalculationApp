package com.company.organizer;

import com.company.organizer.chart.Chart;
import com.company.organizer.chart.ChartData;
import com.company.organizer.chart.bubble.BubbleChart;
import com.company.organizer.chart.bubble.data.BubbleChartData;
import com.company.organizer.chart.bubble.data.ChartSize;
import com.company.organizer.chart.bubble.data.SeriesData;
import com.company.organizer.chart.bubble.data.SourceData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateBubbleChart {

    private static final String SOURCE_SHEET_NAME = "Data";
    private static final String EXCEL_FILE = "bubble.xlsx";

    public static void main(String[] args) throws IOException {
        Workbook workbook = readFile(EXCEL_FILE);

        ChartSize chartSize = new ChartSize.Builder(25, 30)
                .firstColumnIndex(5)
                .firstRowIndex(5)
                .build();

        SeriesData seriesData = new SeriesData();
        seriesData.setName("Project");
        seriesData.setXValue("Seniority per project");
        seriesData.setYValue("PM per Project");
        seriesData.setSize("Employee count");

        SourceData sourceData = new SourceData();
        sourceData.setSeries(seriesData);
        sourceData.setWorkbook(workbook);
        sourceData.setSheetName(SOURCE_SHEET_NAME);

        ChartData chartData = new BubbleChartData(sourceData, chartSize);
        Chart chart = new BubbleChart(chartData);
        chart.fillChart();

        saveChanges(workbook);
    }

    private static XSSFWorkbook readFile(final String fileName) throws IOException {
        try (FileInputStream fis = new FileInputStream(fileName)) {
            return new XSSFWorkbook(fis);
        }
    }

    private static void saveChanges(Workbook wb) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(EXCEL_FILE)) {
            wb.write(fileOut);
        }
    }

}
