package com.company.organizer.chart.bubble;

import com.company.organizer.chart.ChartData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedHashMap;
import java.util.Map;

public class CellReference {

    private static final String CELL_REFERENCE_TEMPLATE = "'%s'!R%dC";
    private final ChartData chartData;
    private Map<String, Integer> columnIndexes;
    private int rowIndex;

    public CellReference(ChartData chartData, int rowIndex) {
        this.chartData = chartData;
        this.rowIndex = rowIndex;
        this.columnIndexes = getSourceColumnIndexes();
    }

    private Map<String, Integer> getSourceColumnIndexes() {
        Sheet dataSheet = getSourceSheet();
        Row row = dataSheet.getRow(0);

        Map<String, Integer> columnIndexes = new LinkedHashMap<>();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            for (String column : chartData.getSourceData().getSeries().getAll()) {
                if (row.getCell(i).getStringCellValue().trim().equals(column)) {
                    columnIndexes.put(column, i + 1);
                    break;
                }
            }
        }
        return columnIndexes;
    }

    private Sheet getSourceSheet() {
        return chartData.getSourceData().getWorkbook().getSheet(chartData.getSourceData().getSheetName());
    }

    public String getReference(String columnName) {
        String template = String.format(CELL_REFERENCE_TEMPLATE, chartData.getSourceData().getSheetName(), rowIndex);
        return template + columnIndexes.get(columnName);
    }

    public String getReference(int columnIndex) {
        String template = String.format(CELL_REFERENCE_TEMPLATE, chartData.getSourceData().getSheetName(), rowIndex);
        return template + columnIndex;
    }

}
