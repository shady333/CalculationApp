package com.company.organizer.core.graphDataSheet;

import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.models.graph.GraphEmp;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

import static com.company.organizer.commons.CommonConst.*;

public class GraphDataSheet {
    int rowEmpCount;
    private BaseExcel baseExcel = new BaseExcel(REVENUE_PATH).openFile();
    private Sheet sheet = baseExcel.createSheet(CHART_SHEET);
    private int BEGIN_ROW_CREATED_SHEET = 1;


    public void writeSheetGraphs(List<GraphEmp> graphEmpList) {
        Row row = sheet.createRow(0);
        row.setHeightInPoints(60);
        for (int i = 0; i < GRAPH_HEADER_NAME.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(GRAPH_HEADER_NAME.get(i));
            cell.setCellStyle(getCellSHeaderStyle());
        }

//        RM
        sheet.setColumnWidth(0, 5000);
//            TA Name
        sheet.setColumnWidth(1, 7000);

        for (int i = 0; i < graphEmpList.size(); i++) {
            Row row1 = createCustomRow();
            GraphEmp graphEmp = graphEmpList.get(i);
//              Project
            Cell cell1 = row1.createCell(0);
            cell1.setCellValue(graphEmp.getProjectName());
            cell1.setCellStyle(getProjectCellStyle());

//              Seniority Per Project
            Cell cell2 = row1.createCell(1);
            cell2.setCellValue(graphEmp.getSeniorityPerProject());
            cell2.setCellStyle(getStandardCellStyle());

//              Sen Count
            Cell cell3 = row1.createCell(2);
            cell3.setCellValue(graphEmp.getProjectPM());
            cell3.setCellStyle(getStandardCellStyle());

            //              Emp Count
            Cell cell4 = row1.createCell(3);
            cell4.setCellValue(graphEmp.getEmpCount());
            cell4.setCellStyle(getStandardCellStyle());


        }

        baseExcel.saveChangesToFile();
    }

    public Row getRow() {
        return sheet.getRow(rowEmpCount++);
    }

    public Row createCustomRow() {
        return sheet.createRow(BEGIN_ROW_CREATED_SHEET++);
    }

    public CellStyle getCountCellStyle() {
        CellStyle style = baseExcel.createCellStyle();

        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        return style;
    }

    public CellStyle getStandardCellStyle() {

        CellStyle style = baseExcel.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    private CellStyle getCellSHeaderStyle() {
        CellStyle style = baseExcel.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setWrapText(true);

        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        return style;
    }

    public CellStyle getProjectCellStyle() {
        CellStyle style = baseExcel.createCellStyle();

        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);

        Font font = baseExcel.getWorkbook().createFont();
        font.setBold(true);

        style.setFont(font);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderTop(BorderStyle.THICK);
        return style;
    }
}
