package com.company.organizer.chart.bubble.data;


import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data from the source sheet to be used in destination chart
 */
public class SourceData {

    private SeriesData series;
    private Workbook workbook;
    private String sheetName;

    public SeriesData getSeries() {
        return series;
    }

    public void setSeries(SeriesData series) {
        this.series = series;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Workbook getWorkbook() {
        return workbook;
    }
}
