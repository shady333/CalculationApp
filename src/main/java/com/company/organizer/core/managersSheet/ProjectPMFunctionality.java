package com.company.organizer.core.managersSheet;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.models.customer.Customers;
import com.company.organizer.models.graph.GraphEmp;
import com.company.organizer.core.base.BaseExcel;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class ProjectPMFunctionality {
    private List<GraphEmp> graphEmpList = new ArrayList<>();
    private BaseExcel baseExcel = new BaseExcel(CommonConst.REVENUE_PATH).openFile();
    private Sheet sheetGet = baseExcel.getSheet(CommonConst.MANAGERS_SHEET_NAME);
    private int BEGIN_ROW_CREATED_SHEET = 0;

    public void setPM(List<Customers> customers) {

        for (int i = 0; i < customers.size(); i++) {
            Customers customer = customers.get(i);

            for (int j = 0; j < customer.getStreamsList().size(); j++) {
                int rowSize = BEGIN_ROW_CREATED_SHEET;
                double sumRevenuePerStream = 0;
                double sumCostPerStream = 0;
                double sumRevenue152perStream = 0;
                double sumLostRevenue = 0;

                int empSize = customer.getStreamsList().get(j).getEmployeesList().size();
                Integer TOTAL_SENIORITY = 0;

                for (int k = 0; k < empSize; k++) {
                    BEGIN_ROW_CREATED_SHEET++;
//                    Revenue
                    double revenue = Double.parseDouble(customer.getStreamsList().get(j).getEmployeesList().get(k).getRevenue());
                    sumRevenuePerStream = sumRevenuePerStream + revenue;

//                   Revenue based on 152
                    String cell = getField(BEGIN_ROW_CREATED_SHEET, 6);
                    double rev152 = Double.parseDouble(cell);
                    sumRevenue152perStream = sumRevenue152perStream + rev152;

                    String cell7 = getField(BEGIN_ROW_CREATED_SHEET, 7);
                    if (cell7.equals("rate mismatch")) {
                    } else {
                        double lost = Double.parseDouble(cell);
                        sumLostRevenue = sumLostRevenue + lost;
                    }

//                    cost
                    String cell1 = getField(BEGIN_ROW_CREATED_SHEET, 11);
                    double cost = Double.parseDouble(cell1);
                    sumCostPerStream = sumCostPerStream + cost;

                    Integer empCount = customer.getStreamsList().get(j).getEmployeesList().get(k).getEmployeeSeniority();
                    TOTAL_SENIORITY = TOTAL_SENIORITY + empCount;

//                                       Cost
                    int numb = BEGIN_ROW_CREATED_SHEET+1;
                    String cellNumb = "K" + numb;
                    sheetGet.getRow(BEGIN_ROW_CREATED_SHEET).createCell(11).setCellFormula("IF(+" + cellNumb + "=T3,W3,IF(" + cellNumb + "=T4,W4,if(" + cellNumb + "=T5,W5,if(" + cellNumb + "=T6,W6,if(" + cellNumb + "=T7,W7)))))");

                }
//                    Project PM
                double projectPM = ((sumRevenuePerStream - sumCostPerStream) / sumRevenuePerStream) * 100;
                double result152 = ((sumRevenue152perStream - sumCostPerStream) / sumRevenue152perStream) * 100;
                Integer empSize1 = customer.getStreamsList().get(j).getEmployeesList().size();

                String rowLabel = customer.getStreamsList().get(j).getRowLabels();

                graphEmpList.add(new GraphEmp(rowLabel, TOTAL_SENIORITY, projectPM, empSize1));

                for (int k = 0; k < empSize1; k++) {
                    rowSize++;
                    Row rowPM = getRow(rowSize);

                    Cell cell17 = rowPM.createCell(17);
                    cell17.setCellStyle(getCountCellStyle());
                    cell17.setCellValue(projectPM);

//                    Customer PM 152
                    Cell cell18 = rowPM.createCell(18);
                    cell18.setCellStyle(getCountCellStyle());
                    cell18.setCellValue(result152);
                }
            }
        }

        baseExcel.saveChangesToFile();
    }

    public List<GraphEmp> getGraphEmpList() {
        return graphEmpList;
    }

    public String getField(int rowNumber, Integer cellNumber) {
        Row row = baseExcel.getSheet(CommonConst.MANAGERS_SHEET_NAME).getRow(rowNumber);
        String string;
        try {
            Cell cellValue = row.getCell(cellNumber);
            switch (cellValue.getCellTypeEnum()) {
                case BOOLEAN:
                    string = String.valueOf(cellValue.getBooleanCellValue());
                    break;
                case NUMERIC:
                    string = String.valueOf(cellValue.getNumericCellValue());
                    break;
                case STRING:
                    string = cellValue.getStringCellValue();
                    break;
                case BLANK:
                    string = "Empty Cell";
                    break;
                case ERROR:
                    string = String.valueOf(cellValue.getErrorCellValue());
                    break;

                case FORMULA:
                    baseExcel.evaluateAllFormulaCells().evaluateInCell(cellValue);
                    string = String.valueOf(cellValue.getNumericCellValue());
                    break;
                default:
                    string = "Wow CELL have unrecognized type";
                    break;
            }
        } catch (NullPointerException e) {
            return String.valueOf(0);
        }

        return string;
    }


    public Row getRow(int rowNumb) {
        return sheetGet.getRow(rowNumb);
    }

    //    Styles

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

    private CellStyle getCellSTotalStyle() {
        CellStyle style = baseExcel.createCellStyle();
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

    public CellStyle getStandardCellStyle() {
        CellStyle style = baseExcel.createCellStyle();

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    public CellStyle getRedCellStyle() {
        CellStyle style = baseExcel.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.RED.getIndex());

        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        return style;
    }

    public CellStyle getYellowCellStyle() {
        CellStyle style = baseExcel.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());

        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        return style;
    }

    public CellStyle getCountCellStyle() {
        CellStyle style = baseExcel.createCellStyle();

        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        DataFormat format = baseExcel.getWorkbook().createDataFormat();

        style.setDataFormat(format.getFormat("#.##"));
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        return style;
    }

}

