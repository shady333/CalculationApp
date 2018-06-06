package com.company.organizer.core.rmSheet;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.models.rm.FullEmployee;
import com.company.organizer.models.rm.RMPersonList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import static java.util.stream.Collectors.toList;

public class RMSheet {
    private int senNumber;
    int rowEmpCount;
    private BaseExcel baseExcel = new BaseExcel(CommonConst.REVENUE_PATH).openFile();
    private Sheet sheet = baseExcel.createSheet(CommonConst.RM_SHEETS);
    private int BEGIN_ROW_CREATED_SHEET = 1;
    private double TOTAL_EMP_COUNT;
    private double TOTAL_SENIORITY;

    public int getSeniority(String value) {
        String normValue;
        if (value.equals("Junior Software Test Automation Engineer")) {
            normValue = CommonConst.JUNIOR;
            senNumber = 1;
        } else if (value.equals("Software Test Automation Engineer")) {
            normValue = CommonConst.INTERMEDIATE;
            senNumber = 2;
        } else if (value.equals("Senior Software Test Automation Engineer") || value.equals("Senior Software Testing Engineer")) {
            normValue = CommonConst.SENIOR;
            senNumber = 3;
        } else if (value.equals("Lead Software Test Automation Engineer") || value.equals("Software Engineering Team Leader")) {
            normValue = CommonConst.LEAD;
            senNumber = 4;
        } else if (value.equals("Software Engineering Manager")) {
            normValue = CommonConst.ABOVE_LEAD;
            senNumber = 5;
        } else {
            senNumber = 0;
            normValue = "doesn't found employee";
        }
        return senNumber;
    }

    public int findPersonSeniority(String title, List<FullEmployee> empList) {
        for (FullEmployee fullEmployee : empList) {
            if (title.equals(fullEmployee.getTitle())) {
                return getSeniority(fullEmployee.getTitle());
            }
        }
        return 0;
    }

    private static List<String> getDistinctRM(List<FullEmployee> employees) {
        List<String> allRMs = new ArrayList<>();
        employees
                .forEach(employee ->
                        allRMs.add(employee.getRm()));

        return allRMs.stream().distinct().collect(toList());
    }

    private static List<FullEmployee> getAllEmpForRM(List<FullEmployee> employees, String id) {
        return employees
                .stream()
                .filter(x -> x.getRm().equalsIgnoreCase(id))
                .collect(toList());
    }

    public List<RMPersonList> processAllEmployees(List<FullEmployee> employeeList) {
        List<String> rms = getDistinctRM(employeeList);
        List<RMPersonList> fullRMsModel = new ArrayList<>();

        employeeList.forEach(emp -> {
            Integer seniority = findPersonSeniority(emp.getTitle(),employeeList);
            emp.setEmployeeSeniority(seniority);
        });
        rms.forEach(rm -> {
            List<FullEmployee> list = getAllEmpForRM(employeeList, rm);
            RMPersonList model = new RMPersonList(rm, list);
            fullRMsModel.add(model);
        });
        return fullRMsModel;
    }

    public void writeSheetRMs(List<RMPersonList> rmPersonLists) {
        Row row = sheet.createRow(0);
        row.setHeightInPoints(60);
        for (int i = 0; i < CommonConst.RM_HEADER_NAME.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(CommonConst.RM_HEADER_NAME.get(i));
            cell.setCellStyle(getCellSHeaderStyle());
        }
//        RM
        sheet.setColumnWidth(0, 5000);
//            TA Name
        sheet.setColumnWidth(1, 7000);

        for (int i = 0; i < rmPersonLists.size(); i++) {
            int rowCount = BEGIN_ROW_CREATED_SHEET;
            rowEmpCount = BEGIN_ROW_CREATED_SHEET;
            int empSenCountPerRM = 0;

            rowCount++;
            for (int j = 0; j < rmPersonLists.get(i).getFullEmployees().size(); j++) {
                Row row1 = createCustomRow();

//              RM
                Cell cell1 = row1.createCell(0);
                cell1.setCellValue(rmPersonLists.get(i).getFullEmployees().get(j).getRm());
                cell1.setCellStyle(getProjectCellStyle());

//              TA Name
                Cell cell2 = row1.createCell(1);
                cell2.setCellValue(rmPersonLists.get(i).getFullEmployees().get(j).getName());
                cell2.setCellStyle(getStandardCellStyle());

//              Sen Count
                Cell cell3 = row1.createCell(2);
                cell3.setCellValue(rmPersonLists.get(i).getFullEmployees().get(j).getEmployeeSeniority());
                cell3.setCellStyle(getStandardCellStyle());

                empSenCountPerRM = empSenCountPerRM + rmPersonLists.get(i).getFullEmployees().get(j).getEmployeeSeniority();
            }


//            Employee count
            Integer EMP_COUNT = rmPersonLists.get(i).getFullEmployees().size();

//                Seniority per rm
            double average = (double) empSenCountPerRM / EMP_COUNT;

//                Second loop for setting total Emp count and Sen per project for merging cells
            for (int k = 0; k < rmPersonLists.get(i).getFullEmployees().size(); k++) {
                Row row1 = getRow();

//                Emp Count
                Cell cell15 = row1.createCell(3);
                cell15.setCellStyle(getCountCellStyle());
                cell15.setCellValue(EMP_COUNT);

//                Average Seniority
                Cell cell16 = row1.createCell(4);
                cell16.setCellStyle(getCountCellStyle());
                cell16.setCellValue(new DecimalFormat("##.##").format(average));

//                Revenue

            }

            // Cell merging
            int size = rowCount + (rmPersonLists.get(i).getFullEmployees().size() - 1);
            if (rowCount != size) {
                sheet.addMergedRegion(CellRangeAddress.valueOf("A" + rowCount + ":A" + size));
                sheet.addMergedRegion(CellRangeAddress.valueOf("D" + rowCount + ":D" + size));
                sheet.addMergedRegion(CellRangeAddress.valueOf("E" + rowCount + ":E" + size));
            }
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
