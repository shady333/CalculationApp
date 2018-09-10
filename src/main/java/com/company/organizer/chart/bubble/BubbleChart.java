package com.company.organizer.chart.bubble;

import com.company.organizer.chart.Chart;
import com.company.organizer.chart.ChartData;
import com.company.organizer.chart.bubble.data.ChartSize;
import com.company.organizer.chart.bubble.data.SeriesData;
import org.apache.commons.lang3.RandomUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.LinkedList;
import java.util.List;
import java.util.Objects;

public class BubbleChart implements Chart {

    private static final String DESTINATION_SHEET_NAME = "Bubble_Chart_";
    private static final int ROW_INDEX_IN_CELL_REFERENCE = 2;
    private static final int AXIS_ID_X = 123456;
    private static final int AXIS_ID_Y = 123457;

    private final Workbook workbook;
    private final Sheet sheet;
    private final ChartData chartData;
    private final SeriesData seriesData;

    private CTChart ctChart;
    private CTPlotArea plotArea;
    private CTBubbleChart bubbleChart;
    private ChartSize chartSize;

    /**
     * @param chartData data to fulfill the chart
     */
    public BubbleChart(ChartData chartData) {
        this.chartData = Objects.requireNonNull(chartData);
        this.workbook = chartData.getSourceData().getWorkbook();
        this.seriesData = chartData.getSourceData().getSeries();
        this.chartSize = chartData.getChartSize();
        this.sheet = createNewSheet();
        this.bubbleChart = createEmptyChart();
    }

    private Sheet createNewSheet() {
        return workbook.createSheet(DESTINATION_SHEET_NAME + RandomUtils.nextInt(1, 999));
    }

    private CTBubbleChart createEmptyChart() {
        Drawing drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = drawChartSize(drawing);

        org.apache.poi.ss.usermodel.Chart chart = drawing.createChart(anchor);
        ctChart = ((XSSFChart) chart).getCTChart();
        plotArea = ctChart.getPlotArea();
        bubbleChart = plotArea.addNewBubbleChart();
        return bubbleChart;
    }

    private ClientAnchor drawChartSize(Drawing drawing) {
        return drawing.createAnchor(0, 0, 0, 0,
                chartSize.getFirstColumnIndex(),
                chartSize.getFirstRowIndex(),
                chartSize.getLastColumnIndex(),
                chartSize.getLastRowIndex());
    }

    public void fillChart() {
        addSeries();
        setIdForAxes();
        addDataLabel();
        addAxis(STAxPos.B, AXIS_ID_X, AXIS_ID_Y);
        addAxis(STAxPos.L, AXIS_ID_Y, AXIS_ID_X);
        addLegend(STLegendPos.B);
    }

    private void addSeries() {
        for (int rowNumber = ROW_INDEX_IN_CELL_REFERENCE; rowNumber <= getSourceSheet().getLastRowNum(); rowNumber++) {
            CellReference cellReference = new CellReference(chartData, rowNumber);
            Series series = new Series(bubbleChart.addNewSer());

            series.setName(cellReference.getReference(seriesData.getName()));
            series.setColor(rowNumber);
            series.setXValue(cellReference.getReference(seriesData.getXValue()));
            series.setYValue(cellReference.getReference(seriesData.getYValue()));
            series.setSize(cellReference.getReference(seriesData.getSize()));

            // compute merged rows
            int rowIndex = rowNumber - 1;
            List<Integer> mergedRows = getRangeOfMergedRows(getSourceSheet(), rowIndex);
            if (mergedRows.size() > 0) {
                rowNumber = mergedRows.get(1) + 1;
            }
        }
    }

    private void setIdForAxes() {
        // telling the Bubble Chart that it has axes and giving them Ids
        bubbleChart.addNewAxId().setVal(AXIS_ID_X);
        bubbleChart.addNewAxId().setVal(AXIS_ID_Y);
    }

    private void addDataLabel() {
        bubbleChart.addNewDLbls().addNewShowBubbleSize().setVal(false);
        bubbleChart.getDLbls().addNewDLblPos().setVal(STDLblPos.Enum.forInt(3));
    }

    private void addAxis(STAxPos.Enum position, int catAxis, int valAxis) {
        CTValAx ctCatAx = plotArea.addNewValAx();
        ctCatAx.addNewAxId().setVal(catAxis); // id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(position);
        ctCatAx.addNewCrossAx().setVal(valAxis); // id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        ctCatAx.addNewMajorGridlines();
    }

    private void addLegend(STLegendPos.Enum position) {
        CTLegend legend = ctChart.addNewLegend();
        legend.addNewLegendPos().setVal(position);
        legend.addNewOverlay().setVal(false);
    }

    private List<Integer> getRangeOfMergedRows(Sheet sheet, int rowIndex) {
        List<Integer> mergedRows = new LinkedList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstRow() <= rowIndex && range.getLastRow() >= rowIndex) {
                mergedRows.add(range.getFirstRow());
                mergedRows.add(range.getLastRow());
            }
        }
        return mergedRows;
    }

    private Sheet getSourceSheet() {
        return workbook.getSheet(chartData.getSourceData().getSheetName());
    }

}
