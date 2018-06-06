package com.company.organizer.core.managersSheet;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.commons.LevelConst;
import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.core.emp.StatusEmp;
import com.company.organizer.models.customer.Customers;
import com.company.organizer.models.rm.FullEmployee;
import com.company.organizer.models.salaryTable.Position;
import org.apache.poi.ss.usermodel.*;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class RevenueSheetWriter {
    StatusEmp statusEmp = new StatusEmp();
    String TITLE = "zero";
    String RATE = "zero";
    String RATE_ZERO = String.valueOf(0.0);
    int rowCount;
    private int LAST_EMP_ROW;
    private double TOTAL_REVENUE;
    private double TOTAL_EMP_COUNT;
    private double TOTAL_SENIORITY;
    private String RATE_MISMATCH = "";
    private BaseExcel baseExcel = new BaseExcel(CommonConst.REVENUE_PATH).openFile();
    private Sheet sheet = baseExcel.createSheet(CommonConst.MANAGERS_SHEET_NAME);

    private int BEGIN_ROW_CREATED_SHEET = 1;

    public void writeLevelsRevenue(List<Position> list) {
        if (list.size() > 0) {
            Row benchRow = sheet.getRow(0);
            for (int i = 19; i < (CommonConst.LEVELS_LIST.size() + 19); i++) {
                Cell cell = benchRow.createCell(i);
                cell.setCellValue(CommonConst.LEVELS_LIST.get(i - 19));
                cell.setCellStyle(getCellSHeaderStyle());
            }

            Row row = sheet.getRow(1);
            getOrCreateCell(row, 22, String.valueOf(LevelConst.DEFAULT_VALUE), getCellSHeaderStyle());


            int LAST_ROW = 0;

            for (int i = 0, j = 2; i < list.size(); i++, j++) {
                Row benchRowForVal = sheet.getRow(j);
                Cell cell19 = getOrCreateCell(benchRowForVal, 19, list.get(i).getLevel(), getCellSHeaderStyle());
                Cell cell20 = getOrCreateCell(benchRowForVal, 20, String.valueOf(list.get(i).getSalary()), getCellSHeaderStyle());
                Cell cell21 = getOrCreateCell(benchRowForVal, 21, String.valueOf(list.get(i).getOverhead()), getCellSHeaderStyle());
                Cell cell22 = getOrCreateCell(benchRowForVal, 22);
                String formula = "U" + (j + 1) + "+V" + (j + 1);
                cell22.setCellFormula(formula);
                cell22.setCellStyle(getCellSHeaderStyle());
                LAST_ROW = j;

            }
            LAST_ROW++;
            Row averageRow = sheet.getRow(LAST_ROW);
            Cell cell19 = getOrCreateCell(averageRow, 19, "Average rate ", getCellSHeaderStyle());
            Cell cell20 = getOrCreateCell(averageRow, 20, String.valueOf(LevelConst.AVERAGE_RATE), getCellSHeaderStyle());

        } else {

            List<Position> positionList = new ArrayList<>();
            positionList.add(new Position(CommonConst.JUNIOR, LevelConst.JUN_SALARY, LevelConst.JUN_OVERHEAD));
            positionList.add(new Position(CommonConst.INTERMEDIATE, LevelConst.INTER_SALARY, LevelConst.INTER_OVERHEAD));
            positionList.add(new Position(CommonConst.SENIOR, LevelConst.SENIOR_SALARY, LevelConst.SENIOR_OVERHEAD));
            positionList.add(new Position(CommonConst.LEAD, LevelConst.LEAD_SALARY, LevelConst.LEAD_OVERHEAD));
            positionList.add(new Position(CommonConst.ABOVE_LEAD, LevelConst.ABOVE_LEAD_SALARY, LevelConst.ABOVE_LEAD_OVERHEAD));

            Row benchRow = sheet.getRow(0);
            for (int i = 19; i < (CommonConst.LEVELS_LIST.size() + 19); i++) {
                Cell cell = benchRow.createCell(i);
                cell.setCellValue(CommonConst.LEVELS_LIST.get(i - 19));
                cell.setCellStyle(getCellSHeaderStyle());
            }

            Row row = sheet.getRow(1);
            getOrCreateCell(row, 22, String.valueOf(LevelConst.DEFAULT_VALUE), getCellSHeaderStyle());

            int LAST_ROW = 0;
            for (int i = 0, j = 2; i < positionList.size(); i++, j++) {

                Row benchRowForVal = sheet.getRow(j);
                Cell cell19 = getOrCreateCell(benchRowForVal, 19, positionList.get(i).getLevel(), getCellSHeaderStyle());
                Cell cell20 = getOrCreateCell(benchRowForVal, 20, String.valueOf(positionList.get(i).getSalary()), getCellSHeaderStyle());
                Cell cell21 = getOrCreateCell(benchRowForVal, 21, String.valueOf(positionList.get(i).getOverhead()), getCellSHeaderStyle());
                Cell cell22 = getOrCreateCell(benchRowForVal, 22);
                String formula = "U" + (j + 1) + "+V" + (j + 1);
                cell22.setCellFormula(formula);
                cell22.setCellStyle(getCellSHeaderStyle());
                LAST_ROW = j;

            }
            LAST_ROW++;
            Row averageRow = sheet.getRow(LAST_ROW);
            Cell cell19 = getOrCreateCell(averageRow, 19, "Average rate ", getCellSHeaderStyle());
            Cell cell20 = getOrCreateCell(averageRow, 20, String.valueOf(LevelConst.AVERAGE_RATE), getCellSHeaderStyle());
        }

        baseExcel.saveChangesToFile();
    }

    public Cell getOrCreateCell(Row row, int numb, String value, CellStyle cellStyle) {
        Cell cell;
        cell = row.createCell(numb);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
        return cell;
    }

    public Cell getOrCreateCell(Row row, int numb) {
        Cell cell;
        try {
            cell = row.getCell(numb);
            if (Objects.isNull(cell)) {
                cell = row.createCell(numb);
            }
            return cell;
        } catch (Exception e) {
            cell = row.createCell(numb);
            return cell;
        }
    }

    public void writeTotal() {
        Row row = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell benchHeader = row.createCell(0);
        benchHeader.setCellValue("TOTAL");
        benchHeader.setCellStyle(getCellSHeaderStyle());

//       revenue
        Cell cell0 = row.createCell(4);
        cell0.setCellValue(TOTAL_REVENUE);
        cell0.setCellStyle(getDollarForFormulaStyle());

//        Revenue based on 152 h
        Cell cell = row.createCell(6);
        cell.setCellFormula("SUM(G2:G" + BEGIN_ROW_CREATED_SHEET + ")");
        cell.setCellStyle(getDollarForFormulaStyle());

//        Lost revenue/0 rate
        Cell cell1 = row.createCell(7);
        cell1.setCellFormula("SUM(H2:H" + BEGIN_ROW_CREATED_SHEET + ")");
        cell1.setCellStyle(getDollarForFormulaStyle());

//        Final revenue
        Cell cell2 = row.createCell(9);
        cell2.setCellFormula("SUM(J2:J" + BEGIN_ROW_CREATED_SHEET + ")");
        cell2.setCellStyle(getDollarForFormulaStyle());

//        Cost
        Cell cell3 = row.createCell(11);
        cell3.setCellFormula("SUM(L2:L" + BEGIN_ROW_CREATED_SHEET + ")");
        cell3.setCellStyle(getDollarForFormulaStyle());

//        Seniority per person
        Cell cell4 = row.createCell(14);
        cell4.setCellValue(new DecimalFormat("##.#").format(TOTAL_SENIORITY / TOTAL_EMP_COUNT));

//        Employee count
        Cell cell5 = row.createCell(15);
        cell5.setCellValue(TOTAL_EMP_COUNT);


        BEGIN_ROW_CREATED_SHEET++;
        baseExcel.saveChangesToFile();
        writeConclusionTotal();
//        baseExcel.saveChangesToFile();
    }

    public void writeConclusionTotal() {
        BEGIN_ROW_CREATED_SHEET = BEGIN_ROW_CREATED_SHEET + 4;

        sheet.setColumnWidth(4, 7000);

        //       revenue
        Row row = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell = row.createCell(4);
        cell.setCellValue("Real Revenue as per report");
        cell.setCellStyle(getCellSTotalStyle());

        Cell cellTotalRev = row.createCell(5);
        cellTotalRev.setCellValue(TOTAL_REVENUE);
        cellTotalRev.setCellStyle(getDollarForFormulaStyle());
        BEGIN_ROW_CREATED_SHEET++;


        //       Revenue based on 152 hours
        Row row1 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell1 = row1.createCell(4);
        cell1.setCellValue("Revenue based on 152 hours");
        cell1.setCellStyle(getCellSTotalStyle());

        String TOTAL_REV_BASED_ON_152 = "SUM(G2:G" + LAST_EMP_ROW + ")";
        Cell cellTotalRevBasedOn152 = row1.createCell(5);
        cellTotalRevBasedOn152.setCellFormula(TOTAL_REV_BASED_ON_152);
        cellTotalRevBasedOn152.setCellStyle(getDollarForFormulaStyle());
        BEGIN_ROW_CREATED_SHEET++;

//        Lost revenue (0 rate)
        Row row2 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell2 = row2.createCell(4);
        cell2.setCellValue("Lost revenue (0 rate)");
        cell2.setCellStyle(getCellSTotalStyle());

        String TOTAL_LOST_REV_ZERO_RATE = "SUM(H2:H" + LAST_EMP_ROW + ")";
        Cell cellTotalLost = row2.createCell(5);
        cellTotalLost.setCellFormula(TOTAL_LOST_REV_ZERO_RATE);
        cellTotalLost.setCellStyle(getDollarForFormulaStyle());
        BEGIN_ROW_CREATED_SHEET++;

//        Ideal Revenue, based on 152 hours and no 0 rates
        Row row3 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell3 = row3.createCell(4);
        cell3.setCellValue("Ideal Revenue, based on 152 hours and no 0 rates");
        cell3.setCellStyle(getCellSTotalStyle());

        String TOTAL_IDEAL_REVENUE = "SUM(J2:J" + LAST_EMP_ROW + ")+SUM(H2:H" + LAST_EMP_ROW + ")";
        Cell cellTotalIdealRev = row3.createCell(5);
        cellTotalIdealRev.setCellFormula(TOTAL_IDEAL_REVENUE);
        cellTotalIdealRev.setCellStyle(getDollarForFormulaStyle());
        BEGIN_ROW_CREATED_SHEET++;

//       Total cost
        Row row4 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell4 = row4.createCell(4);
        cell4.setCellValue("Total cost");
        cell4.setCellStyle(getCellSTotalStyle());

        String TOTAL_COST = "SUM(L2:L" + LAST_EMP_ROW + ")";
        Cell cellTotalCost = row4.createCell(5);
        cellTotalCost.setCellFormula(TOTAL_COST);
        cellTotalCost.setCellStyle(getDollarForFormulaStyle());
        BEGIN_ROW_CREATED_SHEET++;

//       PM Real
        Row row5 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell5 = row5.createCell(4);
        cell5.setCellValue("PM Real");
        cell5.setCellStyle(getCellSTotalStyle());

        String TOTAL_PM_REAL = "(" + TOTAL_REVENUE + "-" + TOTAL_COST + ")" + "/(" + TOTAL_REVENUE + ")" + "*" + 100;
        Cell cellPMreal = row5.createCell(5);
        cellPMreal.setCellFormula(TOTAL_PM_REAL);
        cellPMreal.setCellStyle(getDollarForFormulaStyle());

        BEGIN_ROW_CREATED_SHEET++;

        //       PM Ideal
        Row row6 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell6 = row6.createCell(4);
        cell6.setCellValue("PM Ideal");
        cell6.setCellStyle(getCellSTotalStyle());

        String TOTAL_PM_IDEAL = "(" + TOTAL_REV_BASED_ON_152 + "-" + TOTAL_COST + ")" + "/(" + TOTAL_REV_BASED_ON_152 + ")" + "*" + 100;

        Cell cellPMIdeal = row6.createCell(5);
        cellPMIdeal.setCellFormula(TOTAL_PM_IDEAL);
        cellPMIdeal.setCellStyle(getDollarForFormulaStyle());

        BEGIN_ROW_CREATED_SHEET++;

        String COMPLEX_FORMULA = "(" + TOTAL_REVENUE + "+" + TOTAL_LOST_REV_ZERO_RATE + "-" + TOTAL_COST + ")" + "/" + "(" + TOTAL_REVENUE + "+" + TOTAL_LOST_REV_ZERO_RATE + ")" + "*" + "100";

        //       PM Real +lost
        Row row7 = sheet.createRow(BEGIN_ROW_CREATED_SHEET);
        Cell cell7 = row7.createCell(4);
        cell7.setCellValue("PM Real +lost");
        cell7.setCellStyle(getCellSTotalStyle());

        Cell cellPMRealLost = row7.createCell(5);
        cellPMRealLost.setCellFormula(COMPLEX_FORMULA);
        cellPMRealLost.setCellStyle(getDollarForFormulaStyle());

        BEGIN_ROW_CREATED_SHEET++;

    }

    public void writeBenchList(int rowNumb, List<Customers> customers) {
        List<FullEmployee> list = statusEmp.getBenchList(customers);

        for (FullEmployee fullEmployee : list) {
            BEGIN_ROW_CREATED_SHEET++;
            Row row = sheet.createRow(rowNumb);

            Cell benchHeader = row.createCell(0);
            benchHeader.setCellValue("Bench");
            benchHeader.setCellStyle(getCellSHeaderStyle());

            row.createCell(1).setCellValue(fullEmployee.getName());

//            Title
            Cell cell12 = row.createCell(10);
            cell12.setCellValue(fullEmployee.getTitle());
            cell12.setCellStyle(getStandardCellStyle());

//                   Cost
            String cellNumb = "K" + BEGIN_ROW_CREATED_SHEET;
            row.createCell(11).setCellFormula("IF(+" + cellNumb + "=T3,W3,IF(" + cellNumb + "=T4,W4,if(" + cellNumb + "=T5,W5,if(" + cellNumb + "=T6,W6,if(" + cellNumb + "=T7,W7)))))");

            rowNumb++;
        }
//        for (Map.Entry<EmpTitle, String> entry : benchList.entrySet()) {
//            BEGIN_ROW_CREATED_SHEET++;
//            Row row = sheet.createRow(rowNumb);
//
//            Cell benchHeader = row.createCell(0);
//            benchHeader.setCellValue("Bench");
//            benchHeader.setCellStyle(getCellSHeaderStyle());
//
//            row.createCell(1).setCellValue(entry.getKey().getName());
//
////            Seniority Level
//            Cell cell12 = row.createCell(10);
//            String filter = statusEmp.findPersonTitle(entry.getKey().getStatus());
//            cell12.setCellValue(filter);
//            cell12.setCellStyle(getStandardCellStyle());
//
////                   Cost
//            String cellNumb = "K" + BEGIN_ROW_CREATED_SHEET;
//            row.createCell(11).setCellFormula("IF(+" + cellNumb + "=T3,W3,IF(" + cellNumb + "=T4,W4,if(" + cellNumb + "=T5,W5,if(" + cellNumb + "=T6,W6,if(" + cellNumb + "=T7,W7)))))");
//
//            rowNumb++;
//        }
        LAST_EMP_ROW = BEGIN_ROW_CREATED_SHEET;
        baseExcel.saveChangesToFile();
    }

    public void writeSheetEmp(List<Customers> customers) {

        Row row = sheet.createRow(0);
        row.setHeightInPoints(60);

        for (int i = 0; i < CommonConst.RESULT_COLUMN_HEADER_NAME.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(CommonConst.RESULT_COLUMN_HEADER_NAME.get(i));
            cell.setCellStyle(getCellSHeaderStyle());

//            Project
            sheet.setColumnWidth(0, 3300);
//            TA Name
            sheet.setColumnWidth(1, 5000);
//           Lost Revenue
            sheet.setColumnWidth(7, 5000);
//           Seniority Level
            sheet.setColumnWidth(10, 6000);
        }


        for (int i = 0; i < customers.size(); i++) {
            Customers customer = customers.get(i);
            rowCount = BEGIN_ROW_CREATED_SHEET;
            for (int j = 0; j < customer.getStreamsList().size(); j++) {
                int empCountPerCustomer = 0;
                Cell cell15 = null;
                Cell cell16 = null;
                int empCountSum = 0;

                for (int k = 0; k < customer.getStreamsList().get(j).getEmployeesList().size(); k++) {
                    Row row1 = createCustomRow();

//                    Seniority Level
                    Cell cell10 = row1.createCell(10);
//                   Effective rate
                    Cell cell5 = row1.createCell(5);


                    TITLE = customer.getStreamsList().get(j).getEmployeesList().get(k).getTitle();
                    RATE = customer.getStreamsList().get(j).getEmployeesList().get(k).getEffectiveRate();

                    String RATE1 = null;
                    String TITLE1 = null;
                    try {
                        RATE1 = customer.getStreamsList().get(j).getEmployeesList().get(k + 1).getEffectiveRate();
                        TITLE1 = customer.getStreamsList().get(j).getEmployeesList().get(k + 1).getTitle();
                    } catch (IndexOutOfBoundsException exc) {
                        exc.getStackTrace();
                    }
                    Cell cell7 = row1.createCell(7);
                    cell7.setCellStyle(getStandardCellStyle());
                    if (RATE_MISMATCH.equals("rate mismatch")) {
                        cell7.setCellValue("rate mismatch");
                        RATE_MISMATCH = "";
                    }
                    if (TITLE.equals(CommonConst.DOESN_T_FOUND_EMPLOYEE)) {
                        row1.setRowStyle(getRedCellStyle());
                    } else if (TITLE.equals(TITLE1) && !RATE.equals(RATE1)) {
                        RATE_MISMATCH = "rate mismatch";
                        cell7.setCellValue(RATE_MISMATCH);
                    } else if (RATE.equals(RATE_ZERO)) {
                        row1.setRowStyle(getYellowCellStyle());
                        cell7.setCellFormula("152*U8");
                    }


                    cell10.setCellValue(TITLE);
                    cell10.setCellStyle(getStandardCellStyle());

//                    Effective rate
                    cell5.setCellValue(RATE);
                    cell5.setCellStyle(getStandardCellStyle());


//                    Project
                    Cell cell0 = row1.createCell(0);
                    cell0.setCellValue(customer.getStreamsList().get(j).getRowLabels());
                    cell0.setCellStyle(getStandardCellStyle());

//                    TA Name
                    Cell cell1 = row1.createCell(1);
                    cell1.setCellValue(customer.getStreamsList().get(j).getEmployeesList().get(k).getRowLabels());
                    cell1.setCellStyle(getStandardCellStyle());

//                   Reported Hours
                    Cell cell2 = row1.createCell(2);
                    cell2.setCellValue(customer.getStreamsList().get(j).getEmployeesList().get(k).getReportedHours());
                    cell2.setCellStyle(getStandardCellStyle());

//                   Base Hours
                    Cell cell3 = row1.createCell(3);
                    cell3.setCellValue(CommonConst.BASE_HOURSE);
                    cell3.setCellStyle(getStandardCellStyle());

//                   Revenue
                    Cell cell4 = row1.createCell(4);
                    double revenue = Double.parseDouble(customer.getStreamsList().get(j).getEmployeesList().get(k).getRevenue());
                    TOTAL_REVENUE = TOTAL_REVENUE + revenue;
                    cell4.setCellValue(new DecimalFormat("##").format(revenue));
                    cell4.setCellStyle(getStandardCellStyle());

//                    moved effective rate


//                   Revenue based on 152 h
                    String basedOn152 = "D" + BEGIN_ROW_CREATED_SHEET + "*F" + BEGIN_ROW_CREATED_SHEET;
                    Cell cell6 = row1.createCell(6);
                    cell6.setCellFormula(basedOn152);
                    cell6.setCellStyle(getStandardCellStyle());
//                   Lost revenue/0 rate
//                   Fixed Revenue
//                   Final Revenue

                    Cell cell8 = row1.createCell(8);
                    cell8.setCellStyle(getStandardCellStyle());
                    Cell cell9 = row1.createCell(9);
                    cell9.setCellStyle(getStandardCellStyle());

                    if (customer.getStreamsList().get(j).getEmployeesList().get(k).isFinalRevenue()) {
                        cell8.setCellValue("True");

                        cell9.setCellValue(new DecimalFormat("##").format(revenue));
                    } else {
                        cell9.setCellFormula("G" + BEGIN_ROW_CREATED_SHEET);
                        cell9.setCellStyle(getDollarForFormulaStyle());
                    }

//                    moved Seniority level to top of method

//                   Cost
                    String cellNumb = "K" + BEGIN_ROW_CREATED_SHEET;
                    Cell cell11 = row1.createCell(11);
                    cell11.setCellFormula("IF(+" + cellNumb + "=T3,W3,IF(" + cellNumb + "=T4,W4,if(" + cellNumb + "=T5,W5,if(" + cellNumb + "=T6,W6,if(" + cellNumb + "=T7,W7,(0))))))");
                    cell11.setCellStyle(getStandardCellStyle());

//                   PM Real
                    String E = "E" + BEGIN_ROW_CREATED_SHEET;
                    String L = "L" + BEGIN_ROW_CREATED_SHEET;
                    Cell cell12 = row1.createCell(12);
                    cell12.setCellFormula("((" + E + "-" + L + ")/" + E + ")*100 ");
                    cell12.setCellStyle(getStandardCellStyle());
                    cell12.setCellStyle(getDollarForFormulaStyle());

//                  PM 152 HOURS without 0 rates
                    String G = "G" + BEGIN_ROW_CREATED_SHEET;
//                    String L ="L"+BEGIN_ROW_CREATED_SHEET;
                    Cell cell13 = row1.createCell(13);
                    cell13.setCellFormula("((" + G + "-" + L + ")/" + G + ")*100 ");
                    cell13.setCellStyle(getStandardCellStyle());
                    cell13.setCellStyle(getDollarForFormulaStyle());


//                  Seniority per person
                    Cell cell14 = row1.createCell(14);
                    Integer empCount = customer.getStreamsList().get(j).getEmployeesList().get(k).getEmployeeSeniority();
                    TOTAL_SENIORITY = TOTAL_SENIORITY + empCount;
                    cell14.setCellValue(empCount);
                    cell14.setCellStyle(getStandardCellStyle());


                    empCountSum = empCountSum + customer.getStreamsList().get(j).getEmployeesList().get(k).getEmployeeSeniority();

                }
//                Employee count
                Integer EMP_COUNT = customer.getStreamsList().get(j).getEmployeesList().size();
                TOTAL_EMP_COUNT = TOTAL_EMP_COUNT + EMP_COUNT;

//                Seniority per project
                empCountPerCustomer = empCountPerCustomer + customer.getStreamsList().get(j).getEmployeesList().size();
                double average = (double) empCountSum / empCountPerCustomer;

//                Second loop for setting total Emp count and Sen per project for merging cells
                for (int k = 0; k < customer.getStreamsList().get(j).getEmployeesList().size(); k++) {
                    Row row1 = getRow();

//                  Employee count
                    cell15 = row1.createCell(15);
                    cell15.setCellStyle(getCountCellStyle());
                    cell15.setCellValue(EMP_COUNT);

//                   Seniority per project
                    cell16 = row1.createCell(16);
                    cell16.setCellStyle(getCountCellStyle());


//              Seniority per project
                    cell16.setCellValue(new DecimalFormat("##.##").format(average));

                }
//                int empSize =(customer.getStreamsList().get(j).getEmployeesList().size()-1);
//                if(empSize>1){
//                sheet.addMergedRegion(new CellRangeAddress(lastRow,(lastRow+empSize),0,0));
//                }
            }
        }
        writeBenchList(BEGIN_ROW_CREATED_SHEET, customers);
        writeTotal();
//        CellRangeAddress cellMerge = new CellRangeAddress(1,4,1,1);
//        sheet.addMergedRegion(cellMerge);
//        setBordersToMergedCells(sheet, cellMerge);

        baseExcel.saveChangesToFile();
    }


//    Utils

    public Row createCustomRow() {
        return sheet.createRow(BEGIN_ROW_CREATED_SHEET++);
    }

    public Row getRow() {
        return sheet.getRow(rowCount++);
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
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        return style;
    }

    public CellStyle getStandardCellStyle() {
        if (TITLE.equals(CommonConst.DOESN_T_FOUND_EMPLOYEE)) {
            return getRedCellStyle();
        } else if (RATE.equals(RATE_ZERO)) {
            return getYellowCellStyle();
        }
        CellStyle style = baseExcel.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);

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
        style.setAlignment(HorizontalAlignment.LEFT);

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
        style.setAlignment(HorizontalAlignment.LEFT);

        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        return style;
    }


    public CellStyle getDollarForFormulaStyle() {
        CellStyle style = baseExcel.createCellStyle();
        DataFormat format = baseExcel.getWorkbook().createDataFormat();

        style.setDataFormat(format.getFormat("#"));
        return style;
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
}
