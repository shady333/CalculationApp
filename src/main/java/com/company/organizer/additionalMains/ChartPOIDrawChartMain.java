package com.company.organizer.additionalMains;

import com.aspose.cells.*;
import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.utils.Utils;
import com.company.organizer.core.rev.RevenueParser;

public class ChartPOIDrawChartMain {

    public static void main(String[] args) {
        CommonConst.REVENUE_PATH = CommonConst.OUTPUT_DIRECTORY + "\\" + Utils.getExcelPath();
        RevenueParser parser = new RevenueParser();
        parser.createChartSheet();
        new ChartPOIDrawChartMain().method(CommonConst.REVENUE_PATH);
    }

    public void method(String file){

        Workbook workbook = null;
        try {
            workbook = new Workbook(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        Worksheet sheet = workbook.getWorksheets().get(CommonConst.CHART);
// Generate chart

        int chartIndex = sheet.getCharts().add(ChartType.BUBBLE, 2, 2, 28, 20);
        Chart chart = sheet.getCharts().get(chartIndex);


// Insert series, set bubble size and x values



        SeriesCollection series = chart.getNSeries();
//        Series series1 = series.get()
//        Y - Stream PM   - R2:R88
        series.add("=CreatedSheet!R2:R88", true);
        chart.getNSeries().get(0).setBubbleSizes("=CreatedSheet!P2:P88");

//        X - Seniority per person - O2:O88
        chart.getNSeries().get(0).setXValues("=CreatedSheet!Q2:Q88");
//        Bubble size - Emp count - P2:P88
        chart.getNSeries().get(0).setName("=CreatedSheet!A2:A88");
        chart.getNSeries().get(0).setHasSeriesLines(true);

        DataLabels datalabels = series.get(0).getDataLabels();
//        System.out.println(datalabels.);
//        datalabels.setAutoText(false);
//        datalabels.setText("=CreatedSheet!A2:A88");

// Set a single color the series data points

        chart.getNSeries().setColorVaried(true);
        try {
            workbook.save(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
